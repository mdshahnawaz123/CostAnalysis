using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using CostAnalysis.Extension; // DataLab
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CostAnalysis.Command
{
    internal class AssignAbsHandler : IExternalEventHandler
    {
        // Inputs to be set by the UI before calling ExternalEvent.Raise()
        public IList<ElementId> TargetRoomIds { get; set; } = new List<ElementId>();
        public string RoomNumber { get; set; } = "";
        public string RoomLevelName { get; set; } = "";
        public string RoomAssetCode { get; set; } = "";
        public string SharedParameterFileNameHint { get; set; } = "ExpoCity_SharedParameters_for_ABS.txt";

        // Asset Registry for validation
        public HashSet<string> AssetRegistry { get; set; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public void Execute(UIApplication uiapp)
        {
            if (uiapp == null) { TaskDialog.Show("Assign ABS", "UIApplication is null."); return; }

            UIDocument udoc = uiapp.ActiveUIDocument;
            if (udoc == null) { TaskDialog.Show("Assign ABS", "Active document not available."); return; }

            Document doc = udoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = doc.Application;

            // Open or create shared parameter file
            DefinitionFile defFile = app.OpenSharedParameterFile();
            if (defFile == null)
            {
                try
                {
                    string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    string path = Path.Combine(docs, SharedParameterFileNameHint);

                    if (!File.Exists(path))
                        File.WriteAllText(path, "# ExpoCity shared parameters created by AssignAbsHandler\n");

                    app.SharedParametersFilename = path;
                    defFile = app.OpenSharedParameterFile();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Assign ABS", $"Could not create/open shared parameter file:\n{ex.Message}");
                    return;
                }
            }

            if (defFile == null)
            {
                TaskDialog.Show("Assign ABS", "Unable to open or create a Shared Parameters file.");
                return;
            }

            // Ensure a group exists (we put defs in "ABS" group)
            string groupName = "ABS";
            DefinitionGroup defGroup = defFile.Groups.get_Item(groupName) ?? defFile.Groups.Create(groupName);

            // Collect existing definitions across all groups
            var existingDefs = new Dictionary<string, ExternalDefinition>(StringComparer.OrdinalIgnoreCase);
            foreach (DefinitionGroup g in defFile.Groups)
            {
                foreach (Definition d in g.Definitions)
                {
                    if (d is ExternalDefinition ed && !existingDefs.ContainsKey(ed.Name))
                        existingDefs.Add(ed.Name, ed);
                }
            }

            // Create missing definitions in our group
            var createdDefs = new List<ExternalDefinition>();
            var failedToCreate = new List<string>();

            foreach (var pname in DataLab.AbsParamNames)
            {
                if (existingDefs.ContainsKey(pname)) continue;

                // Use ForgeTypeId instead of ParameterType (Revit 2024+)
                ForgeTypeId paramType = SpecTypeId.String.Text;
                if (pname.Equals("ECD_ABS_L6_Quantity", StringComparison.OrdinalIgnoreCase))
                    paramType = SpecTypeId.Int.Integer;

                var opts = new ExternalDefinitionCreationOptions(pname, paramType) { Visible = true };
                try
                {
                    var newDef = defGroup.Definitions.Create(opts) as ExternalDefinition;
                    if (newDef != null)
                    {
                        existingDefs[newDef.Name] = newDef;
                        createdDefs.Add(newDef);
                    }
                    else failedToCreate.Add(pname);
                }
                catch
                {
                    failedToCreate.Add(pname);
                }
            }

            // Build CategorySet (all model categories by default)
            CategorySet cats = app.Create.NewCategorySet();
            foreach (Category c in doc.Settings.Categories)
            {
                if (c != null && c.CategoryType == CategoryType.Model)
                    cats.Insert(c);
            }

            var bindingMap = doc.ParameterBindings;
            var failedToBind = new List<string>();
            var failedToSet = new List<int>();
            var processedElements = 0;
            var skippedDueToRegistry = 0;
            var registryValidated = 0;

            using (var tx = new Transaction(doc, "Bind ABS Shared Parameters and Assign Values"))
            {
                if (tx.Start() != TransactionStatus.Started)
                {
                    TaskDialog.Show("Assign ABS", "Could not start transaction.");
                    return;
                }

                // Bind all existing defs (if not already)
                foreach (var kv in existingDefs)
                {
                    var def = kv.Value;
                    try
                    {
                        if (!bindingMap.Contains(def))
                        {
                            InstanceBinding instBind = app.Create.NewInstanceBinding(cats);
                            bool inserted = bindingMap.Insert(def, instBind, BuiltInParameterGroup.PG_DATA);
                            if (!inserted) failedToBind.Add(def.Name);
                        }
                    }
                    catch
                    {
                        failedToBind.Add(def.Name);
                    }
                }

                // Process rooms and their elements
                foreach (var roomId in TargetRoomIds)
                {
                    try
                    {
                        Element roomElement = doc.GetElement(roomId);
                        if (!(roomElement is Room room)) continue;

                        // Get room's asset code
                        string roomAssetFromRoom = GetRoomAssetCode(room);

                        // Validate: room asset must match the provided asset code
                        if (!string.Equals(roomAssetFromRoom, RoomAssetCode, StringComparison.OrdinalIgnoreCase))
                        {
                            TaskDialog.Show("Asset Mismatch",
                                $"Room '{room.Number}' has asset code '{roomAssetFromRoom}' " +
                                $"which doesn't match provided asset '{RoomAssetCode}'.\n\nSkipping this room.");
                            continue;
                        }

                        // Get elements inside this room
                        var elementsInRoom = GetElementsInRoom(doc, room);

                        foreach (var element in elementsInRoom)
                        {
                            try
                            {
                                // Skip rooms, spaces, areas themselves
                                if (element is Room || element is Space || element is Area) continue;

                                // Get element's current asset code (if any)
                                string elementAsset = GetElementAssetCode(element);

                                // Check Asset Registry validation
                                bool isValidAsset = true;
                                if (AssetRegistry != null && AssetRegistry.Count > 0)
                                {
                                    // If element already has an asset code, validate it against registry
                                    if (!string.IsNullOrEmpty(elementAsset))
                                    {
                                        if (!AssetRegistry.Contains(elementAsset))
                                        {
                                            // Element has asset code but it's not in registry - skip it
                                            skippedDueToRegistry++;
                                            continue;
                                        }
                                        registryValidated++;
                                    }
                                    // If element has no asset code, we'll assign the room's asset
                                    // and assume it's valid since the room asset was already validated
                                }

                                // Assign ABS parameters from room data
                                Parameter pAsset = element.LookupParameter("ECD_ABS_L1_Asset");
                                if (pAsset != null && !pAsset.IsReadOnly)
                                    SetParameterValue(pAsset, RoomAssetCode ?? "");

                                Parameter pLevel = element.LookupParameter("ECD_ABS_L2_Level");
                                if (pLevel != null && !pLevel.IsReadOnly)
                                    SetParameterValue(pLevel, RoomLevelName ?? "");

                                Parameter pRoom = element.LookupParameter("ECD_ABS_L3_Room");
                                if (pRoom != null && !pRoom.IsReadOnly)
                                    SetParameterValue(pRoom, RoomNumber ?? "");

                                Parameter pQty = element.LookupParameter("ECD_ABS_L6_Quantity");
                                if (pQty != null && !pQty.IsReadOnly)
                                    SetParameterValue(pQty, 1);

                                processedElements++;
                            }
                            catch
                            {
                                failedToSet.Add(element.Id.IntegerValue);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Error", $"Failed to process room {roomId}: {ex.Message}");
                    }
                }

                tx.Commit();
            }

            // Build summary
            var sb = new StringBuilder();
            sb.AppendLine($"Assign ABS finished.");
            sb.AppendLine($"✓ Processed {processedElements} elements from {TargetRoomIds.Count} room(s).");

            if (AssetRegistry != null && AssetRegistry.Count > 0)
            {
                sb.AppendLine($"✓ Registry validated: {registryValidated} elements");
                if (skippedDueToRegistry > 0)
                    sb.AppendLine($"⚠ Skipped {skippedDueToRegistry} elements (asset not in registry)");
            }

            if (createdDefs.Any())
            {
                sb.AppendLine(); sb.AppendLine("Created definitions:");
                sb.AppendLine(string.Join(", ", createdDefs.Select(d => d.Name)));
            }
            if (failedToCreate.Any())
            {
                sb.AppendLine(); sb.AppendLine("Failed to create:");
                sb.AppendLine(string.Join(", ", failedToCreate));
            }
            if (failedToBind.Any())
            {
                sb.AppendLine(); sb.AppendLine("Failed to bind:");
                sb.AppendLine(string.Join(", ", failedToBind));
            }
            if (failedToSet.Any())
            {
                sb.AppendLine(); sb.AppendLine($"Failed to set on {failedToSet.Count} elements (IDs):");
                sb.AppendLine(string.Join(", ", failedToSet));
            }

            TaskDialog.Show("Assign ABS", sb.ToString());
        }

        /// <summary>
        /// Gets the asset code from a room's ECD_ABS_L1_Asset parameter
        /// </summary>
        private string GetRoomAssetCode(Room room)
        {
            try
            {
                var p = room?.LookupParameter("ECD_ABS_L1_Asset");
                if (p != null && p.HasValue)
                {
                    switch (p.StorageType)
                    {
                        case StorageType.String: return p.AsString() ?? "";
                        case StorageType.Integer: return p.AsInteger().ToString();
                        case StorageType.Double: return p.AsDouble().ToString();
                    }
                }
            }
            catch { }
            return "";
        }

        /// <summary>
        /// Gets the asset code from an element's ECD_ABS_L1_Asset parameter
        /// </summary>
        private string GetElementAssetCode(Element element)
        {
            try
            {
                var p = element?.LookupParameter("ECD_ABS_L1_Asset");
                if (p != null && p.HasValue)
                {
                    switch (p.StorageType)
                    {
                        case StorageType.String: return p.AsString() ?? "";
                        case StorageType.Integer: return p.AsInteger().ToString();
                        case StorageType.Double: return p.AsDouble().ToString();
                    }
                }
            }
            catch { }
            return "";
        }

        /// <summary>
        /// Gets all elements located inside the given room
        /// </summary>
        private List<Element> GetElementsInRoom(Document doc, Room room)
        {
            var elementsInRoom = new List<Element>();

            try
            {
                // Get the room's phase
                Phase roomPhase = doc.GetElement(room.CreatedPhaseId) as Phase;
                if (roomPhase == null)
                    roomPhase = doc.Phases.get_Item(doc.Phases.Size - 1) as Phase; // Use last phase as default

                // Method 1: Use spatial relationship with phase
                var levelElements = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Where(e => e.LevelId == room.LevelId)
                    .Cast<FamilyInstance>();

                foreach (var fi in levelElements)
                {
                    try
                    {
                        if (fi.Location is LocationPoint loc)
                        {
                            Room elementRoom = doc.GetRoomAtPoint(loc.Point, roomPhase);
                            if (elementRoom != null && elementRoom.Id == room.Id)
                            {
                                if (!elementsInRoom.Any(e => e.Id == fi.Id))
                                    elementsInRoom.Add(fi);
                            }
                        }
                    }
                    catch { }
                }

                // Method 2: Check all elements with location points
                var allElements = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .Where(e => e.Location is LocationPoint && e.LevelId == room.LevelId);

                foreach (var element in allElements)
                {
                    try
                    {
                        if (element.Location is LocationPoint locPt)
                        {
                            Room elementRoom = doc.GetRoomAtPoint(locPt.Point, roomPhase);
                            if (elementRoom != null && elementRoom.Id == room.Id)
                            {
                                if (!elementsInRoom.Any(e => e.Id == element.Id))
                                    elementsInRoom.Add(element);
                            }
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Debug", $"Error finding elements in room: {ex.Message}");
            }

            return elementsInRoom;
        }

        public string GetName() => "AssignAbsHandler";

        public void SetParameterValue(Parameter p, object value)
        {
            if (p == null) return;
            if (p.IsReadOnly) return;

            switch (p.StorageType)
            {
                case StorageType.String:
                    p.Set(value?.ToString() ?? "");
                    break;
                case StorageType.Integer:
                    if (value is int vi) p.Set(vi);
                    else if (int.TryParse(value?.ToString(), out int iv)) p.Set(iv);
                    else p.Set(0);
                    break;
                case StorageType.Double:
                    if (value is double dv) p.Set(dv);
                    else if (double.TryParse(value?.ToString(), out double dd)) p.Set(dd);
                    else p.Set(0.0);
                    break;
                case StorageType.ElementId:
                    break;
            }
        }
    }
}