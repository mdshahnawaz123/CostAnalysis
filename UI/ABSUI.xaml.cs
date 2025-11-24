using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using CostAnalysis.Command;    // AssignAbsHandler
using CostAnalysis.Extension;  // DataLab

namespace CostAnalysis.UI
{
    public partial class ABSUI : Window
    {
        UIDocument uidoc;
        Document doc;
        private List<LinkedRoomItem> _allRooms = new List<LinkedRoomItem>();

        private AssignAbsHandler _assignAbsHandler;
        private ExternalEvent _assignAbsEvent;

        public ABSUI(Document doc, UIDocument uidoc)
        {
            InitializeComponent();
            this.doc = doc ?? throw new ArgumentNullException(nameof(doc));
            this.uidoc = uidoc ?? throw new ArgumentNullException(nameof(uidoc));

            _assignAbsHandler = new AssignAbsHandler();
            _assignAbsEvent = ExternalEvent.Create(_assignAbsHandler);

            DataContext = this;

            LoadRooms();

            if (RoomsFilterText != null)
                RoomsFilterText.TextChanged += (s, e) => ApplyRoomFilter(RoomsFilterText.Text);
        }

        // -------- UI Events --------
        public void RoomsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                ElementsList.ItemsSource = null;

                var sel = RoomsList.SelectedItem as LinkedRoomItem;
                if (sel == null || sel.Room == null) return;

                // Build room solid in the room's source document
                var roomSolidSource = GetRoomSolidFromBoundaries(sel.Room);
                if (roomSolidSource == null)
                {
                    TaskDialog.Show("ABS",
                        "Could not compute a 3D solid for the selected room.\nEnsure the room is properly enclosed in its phase.");
                    return;
                }

                // Transform into HOST coordinates (Identity for host rooms)
                Solid hostSolid = roomSolidSource;
                if (sel.LinkToHost != null && !sel.LinkToHost.IsIdentity)
                    hostSolid = SolidUtils.CreateTransformed(roomSolidSource, sel.LinkToHost);

                // Intersect with HOST elements (only the active document)
                var intersectsFilter = new ElementIntersectsSolidFilter(hostSolid);

                var elems = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .WherePasses(intersectsFilter)
                    .Where(e2 => e2.Category != null && e2.Category.CategoryType == CategoryType.Model)
                    .OrderBy(e2 => e2.Category?.Name)
                    .ThenBy(e2 => e2.Name)
                    .ToList();

                ElementsList.ItemsSource = elems;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("ABS", $"Room load failed.\n\n{ex.Message}");
            }
        }

        public void Btn_Highlight(object sender, RoutedEventArgs e)
        {
            try
            {
                var ids = ElementsList.SelectedItems.Cast<Element>().Select(el => el.Id).ToList();
                if (!ids.Any()) return;

                uidoc.Selection.SetElementIds(ids);
                uidoc.ShowElements(ids);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("ABS", $"Highlight failed.\n\n{ex.Message}");
            }
        }

        // Btn_ABS: Validate room asset, check registry, and assign to elements inside room
        public void Btn_ABS(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. Get selected room
                var selRoom = RoomsList.SelectedItem as LinkedRoomItem;
                if (selRoom == null)
                {
                    TaskDialog.Show("ABS", "Please select a room first.");
                    return;
                }

                var room = selRoom.Room;
                if (room == null)
                {
                    TaskDialog.Show("ABS", "Invalid room selected.");
                    return;
                }

                // 2. Get room data
                string roomNumber = room.Number ?? room.Name ?? "";
                string roomLevel = room.Level?.Name ?? GetLevelName(room);

                // 3. Get room's asset code from its parameter
                string roomAssetCode = "";
                try
                {
                    // Try different parameter name variations
                    string[] paramNames = new[]
                    {
                        "ECD_ABS_L1_Asset",
                        "(01)ECD_ABS_L1_Asset",
                        "01)ECD_ABS_L1_Asset",
                        "ECD_ABS_L1_WBS_Plot_Code",
                        "(02)ECD_ABS_L1_WBS_Plot_Code"
                    };

                    foreach (var paramName in paramNames)
                    {
                        var rp = room.LookupParameter(paramName);
                        if (rp != null && rp.HasValue)
                        {
                            if (rp.StorageType == StorageType.String)
                            {
                                roomAssetCode = rp.AsString() ?? "";
                                if (!string.IsNullOrWhiteSpace(roomAssetCode)) break;
                            }
                            else if (rp.StorageType == StorageType.Integer)
                            {
                                roomAssetCode = rp.AsInteger().ToString();
                                if (!string.IsNullOrWhiteSpace(roomAssetCode)) break;
                            }
                            else if (rp.StorageType == StorageType.Double)
                            {
                                roomAssetCode = rp.AsDouble().ToString();
                                if (!string.IsNullOrWhiteSpace(roomAssetCode)) break;
                            }
                        }
                    }
                }
                catch { }

                // 4. Get typed asset code from textbox
                string typedAssetCode = AssetCodeText?.Text?.Trim() ?? "";

                // 5. Validate: Room asset code must match textbox
                if (string.IsNullOrEmpty(roomAssetCode))
                {
                    TaskDialog.Show("ABS",
                        $"Room '{roomNumber}' has no asset code assigned.\n\n" +
                        "Please assign an asset code to the room first.");
                    return;
                }

                if (string.IsNullOrEmpty(typedAssetCode))
                {
                    TaskDialog.Show("ABS", "Please enter an asset code in the textbox.");
                    return;
                }

                if (!string.Equals(roomAssetCode, typedAssetCode, StringComparison.OrdinalIgnoreCase))
                {
                    TaskDialog.Show("ABS - Asset Mismatch",
                        $"Asset code mismatch!\n\n" +
                        $"Room asset code: {roomAssetCode}\n" +
                        $"Typed asset code: {typedAssetCode}\n\n" +
                        "These must match to proceed. Please check and correct.");
                    return;
                }

                // 6. Load Asset Registry (optional - if you have an asset registry file)
                var assetRegistry = LoadAssetRegistry();

                // 7. Validate asset code against registry (if registry exists)
                if (assetRegistry != null && assetRegistry.Count > 0)
                {
                    if (!assetRegistry.Contains(typedAssetCode))
                    {
                        var result = TaskDialog.Show("ABS - Registry Check",
                            $"Warning: Asset code '{typedAssetCode}' is not found in the asset registry.\n\n" +
                            "Do you want to proceed anyway?",
                            TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                        if (result != TaskDialogResult.Yes)
                            return;
                    }
                }

                // 8. All validations passed - prepare handler
                _assignAbsHandler.TargetRoomIds = new List<ElementId> { room.Id };
                _assignAbsHandler.RoomNumber = roomNumber;
                _assignAbsHandler.RoomLevelName = roomLevel;
                _assignAbsHandler.RoomAssetCode = typedAssetCode;
                _assignAbsHandler.AssetRegistry = assetRegistry;
                _assignAbsHandler.SharedParameterFileNameHint = "ExpoCity_SharedParameters_for_ABS.txt";

                // 9. Raise external event to process
                _assignAbsEvent.Raise();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("ABS Error", $"An error occurred:\n\n{ex.Message}");
            }
        }

        /// <summary>
        /// Loads asset registry from a file or returns an empty set
        /// Modify this method to load from your actual asset registry source
        /// </summary>
        private HashSet<string> LoadAssetRegistry()
        {
            var registry = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // OPTION 1: Load from a CSV file
                string registryPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "AssetRegistry.csv");

                if (File.Exists(registryPath))
                {
                    var lines = File.ReadAllLines(registryPath);
                    foreach (var line in lines.Skip(1)) // Skip header
                    {
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            var parts = line.Split(',');
                            if (parts.Length > 0)
                            {
                                var assetCode = parts[0].Trim().Trim('"');
                                if (!string.IsNullOrEmpty(assetCode))
                                    registry.Add(assetCode);
                            }
                        }
                    }
                }

                // OPTION 2: Load from Revit project parameter or shared parameter
                // You can query all unique asset codes from existing elements in the project
                /*
                var collector = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .ToElements();
                
                foreach (var el in collector)
                {
                    var p = el.LookupParameter("ECD_ABS_L1_Asset");
                    if (p != null && p.HasValue && p.StorageType == StorageType.String)
                    {
                        var code = p.AsString();
                        if (!string.IsNullOrEmpty(code))
                            registry.Add(code);
                    }
                }
                */

                // OPTION 3: Hardcode valid asset codes for testing
                /*
                registry.Add("ASSET_001");
                registry.Add("ASSET_002");
                registry.Add("EXPO_CITY_123");
                */
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Registry Load", $"Could not load asset registry:\n{ex.Message}\n\nProceeding without registry validation.");
            }

            return registry;
        }

        public void Btn_Export(object sender, RoutedEventArgs e)
        {
            try
            {
                var selRoom = RoomsList.SelectedItem as LinkedRoomItem;
                if (selRoom == null)
                {
                    TaskDialog.Show("Export", "Select a room first.");
                    return;
                }

                var records = BuildRecordsForRoom(selRoom);
                if (records == null || records.Count == 0)
                {
                    TaskDialog.Show("Export", "No elements to export for this room.");
                    return;
                }

                var sfd = new SaveFileDialog
                {
                    Title = "Export elements to CSV",
                    Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                    FileName = $"Room_{selRoom.Number}_{selRoom.SourceTag}.csv"
                };

                bool? result = sfd.ShowDialog();
                if (result != true) return;

                using (var sw = new StreamWriter(sfd.FileName, false, Encoding.UTF8))
                {
                    sw.WriteLine(ElementRecord.CsvHeader());
                    foreach (var r in records)
                        sw.WriteLine(r.ToCsvLine());
                }

                TaskDialog.Show("Export", $"Exported {records.Count} records to:\n{sfd.FileName}");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Export", $"Export failed.\n\n{ex.Message}");
            }
        }

        // ------ Filtering / Loading -------------------------------------------------

        public void ApplyRoomFilter(string term)
        {
            if (string.IsNullOrWhiteSpace(term))
            {
                RoomsList.ItemsSource = _allRooms;
                return;
            }

            term = term.Trim();
            var filtered = _allRooms.Where(r =>
                    (r.Number ?? "").IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    (r.Name ?? "").IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    (r.SourceTag ?? "").IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            RoomsList.ItemsSource = filtered;
        }

        public void LoadRooms()
        {
            var items = new List<LinkedRoomItem>();

            // Host rooms
            var hostRooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Cast<Room>()
                .Where(r => r != null && r.Location != null)
                .ToList();

            foreach (var r in hostRooms) items.Add(LinkedRoomItem.FromHost(r, doc));

            // Linked rooms
            var linkInstances = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_RvtLinks)
                .WhereElementIsNotElementType()
                .OfType<RevitLinkInstance>()
                .ToList();

            foreach (var linkInst in linkInstances)
            {
                var linkDoc = linkInst.GetLinkDocument();
                if (linkDoc == null) continue; // unloaded or not available

                Transform linkToHost = linkInst.GetTotalTransform();

                var linkRooms = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()
                    .Cast<Room>()
                    .Where(r => r != null && r.Location != null)
                    .ToList();

                foreach (var r in linkRooms)
                    items.Add(LinkedRoomItem.FromLink(r, linkDoc, linkToHost, linkInst.Name));
            }

            _allRooms = items
                .OrderBy(i => i.IsFromLink)
                .ThenBy(i => i.SourceTag)
                .ThenBy(i => i.Number)
                .ThenBy(i => i.Name)
                .ToList();

            RoomsList.ItemsSource = _allRooms;
            RoomsList.DisplayMemberPath = nameof(LinkedRoomItem.DisplayText);
        }

        private void BuildCategoryTree()
        {
            // optional: populate the CategoriesTree if desired
        }

        // ------ Geometry -----------------------------------------------------------

        private Solid GetRoomSolidFromBoundaries(Room room)
        {
            if (room == null) return null;

            var sebo = new SpatialElementBoundaryOptions
            {
                SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
            };

            IList<IList<BoundarySegment>> loops;
            try
            {
                loops = room.GetBoundarySegments(sebo);
            }
            catch
            {
                return null;
            }

            if (loops == null || loops.Count == 0) return null;

            double height = room.UnboundedHeight;
            if (height <= 1e-6 || double.IsNaN(height))
            {
                const double mmPerFoot = 304.8;
                height = 3000.0 / mmPerFoot;
            }

            var curveLoops = new List<CurveLoop>();
            foreach (var segs in loops)
            {
                var cl = new CurveLoop();
                foreach (var seg in segs)
                {
                    var c = seg?.GetCurve();
                    if (c != null) cl.Append(c);
                }
                if (cl.Count() >= 3) curveLoops.Add(cl);
            }

            if (curveLoops.Count == 0) return null;

            try
            {
                var extruded = GeometryCreationUtilities.CreateExtrusionGeometry(curveLoops, XYZ.BasisZ, height);
                if (extruded != null && extruded.Volume > 1e-9) return extruded;
            }
            catch { }

            return null;
        }

        // -------------------- Element record helpers used for export ------------------

        internal class ElementRecord
        {
            public int Id { get; set; }
            public string Category { get; set; }
            public string Name { get; set; }
            public string TypeName { get; set; }
            public string Level { get; set; }
            public double Volume { get; set; }        // Revit internal units
            public double Area { get; set; }          // Revit internal units
            public string ABSCode { get; set; }       // stored ABS code (if any)
            public string Source { get; set; }        // Host or Link name

            public string ToCsvLine()
            {
                string Esc(string s) => s?.Replace("\"", "\"\"") ?? "";
                return $"\"{Id}\",\"{Esc(Category)}\",\"{Esc(Name)}\",\"{Esc(TypeName)}\",\"{Esc(Level)}\",\"{Volume}\",\"{Area}\",\"{Esc(ABSCode)}\",\"{Esc(Source)}\"";
            }

            public static string CsvHeader()
            {
                return "\"ElementId\",\"Category\",\"Name\",\"TypeName\",\"Level\",\"Volume\",\"Area\",\"ABSCode\",\"Source\"";
            }
        }

        private List<ElementRecord> BuildRecordsForRoom(LinkedRoomItem sel)
        {
            var records = new List<ElementRecord>();
            if (sel == null) return records;

            var roomSolidSource = GetRoomSolidFromBoundaries(sel.Room);
            if (roomSolidSource == null) return records;

            Solid hostSolid = roomSolidSource;
            if (sel.LinkToHost != null && !sel.LinkToHost.IsIdentity)
                hostSolid = SolidUtils.CreateTransformed(roomSolidSource, sel.LinkToHost);

            var intersectsFilter = new ElementIntersectsSolidFilter(hostSolid);

            var elems = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType()
                        .WherePasses(intersectsFilter)
                        .Where(e2 => e2.Category != null && e2.Category.CategoryType == CategoryType.Model)
                        .ToList();

            foreach (var el in elems)
            {
                try
                {
                    records.Add(new ElementRecord
                    {
                        Id = el.Id.IntegerValue,
                        Category = el.Category?.Name,
                        Name = el.Name,
                        TypeName = GetTypeName(el),
                        Level = GetLevelName(el),
                        Volume = GetElementVolume(el, doc),
                        Area = GetElementArea(el),
                        ABSCode = GetStringParameter(el, DataLab.AbsParamNames.FirstOrDefault() ?? "ECD_ABS_L1_Asset"),
                        Source = sel.IsFromLink ? sel.LinkInstanceName : "Host"
                    });
                }
                catch { }
            }

            return records;
        }

        private string GetTypeName(Element e)
        {
            try
            {
                ElementId typeId = e.GetTypeId();
                if (typeId != null && typeId.IntegerValue != -1)
                {
                    var t = doc.GetElement(typeId);
                    if (t != null) return t.Name;
                }
            }
            catch { }
            return "";
        }

        private string GetLevelName(Element e)
        {
            try
            {
                var p = e.get_Parameter(BuiltInParameter.LEVEL_PARAM);
                if (p != null && p.HasValue)
                {
                    var id = p.AsElementId();
                    if (id != null && id?.IntegerValue != -1)
                    {
                        var level = doc.GetElement(id) as Level;
                        if (level != null) return level.Name;
                    }
                }

                var lid = e.LevelId;
                if (lid != null && lid?.IntegerValue != -1)
                {
                    var level = doc.GetElement(lid) as Level;
                    if (level != null) return level.Name;
                }
            }
            catch { }

            return "";
        }

        private double GetElementArea(Element e)
        {
            try
            {
                var p = e.LookupParameter("Area") ?? e.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                if (p != null && p.StorageType == StorageType.Double && p.HasValue)
                    return p.AsDouble();
            }
            catch { }
            return 0.0;
        }

        private static double GetElementVolume(Element e, Document doc)
        {
            try
            {
                var p = e.LookupParameter("Volume") ?? e.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                if (p != null && p.HasValue && p.StorageType == StorageType.Double)
                {
                    return p.AsDouble();
                }
            }
            catch { }

            try
            {
                double sumVol = 0.0;
                var options = new Options { ComputeReferences = false, DetailLevel = ViewDetailLevel.Fine };
                var geomElem = e.get_Geometry(options);
                if (geomElem != null)
                {
                    foreach (GeometryObject gObj in geomElem)
                    {
                        if (gObj is Solid s && s.Volume > 0)
                        {
                            sumVol += s.Volume;
                        }
                        else if (gObj is GeometryInstance gi)
                        {
                            var instGeom = gi.GetInstanceGeometry();
                            foreach (GeometryObject instObj in instGeom)
                            {
                                if (instObj is Solid s2 && s2.Volume > 0) sumVol += s2.Volume;
                            }
                        }
                    }
                }

                return sumVol;
            }
            catch
            {
                return 0.0;
            }
        }

        private string GetStringParameter(Element el, string paramName)
        {
            try
            {
                var p = el.LookupParameter(paramName);
                if (p != null && p.HasValue)
                {
                    switch (p.StorageType)
                    {
                        case StorageType.String: return p.AsString();
                        case StorageType.Double: return p.AsDouble().ToString();
                        case StorageType.Integer: return p.AsInteger().ToString();
                        default: return "";
                    }
                }
            }
            catch { }
            return "";
        }
    }

    internal class LinkedRoomItem
    {
        public Room Room { get; private set; }
        public Document SourceDoc { get; private set; }
        public Transform LinkToHost { get; private set; } // Identity for host
        public bool IsFromLink { get; private set; }
        public string LinkInstanceName { get; private set; }

        public string Number => Room?.Number;
        public string Name => Room?.Name;

        public string SourceTag => IsFromLink ? (string.IsNullOrEmpty(LinkInstanceName) ? "Link" : LinkInstanceName) : "Host";

        public string DisplayText => $"{Number} - {Name}  [{SourceTag}]";

        private LinkedRoomItem() { }

        public static LinkedRoomItem FromHost(Room room, Document hostDoc)
        {
            return new LinkedRoomItem { Room = room, SourceDoc = hostDoc, LinkToHost = Transform.Identity, IsFromLink = false, LinkInstanceName = "Host" };
        }

        public static LinkedRoomItem FromLink(Room room, Document linkDoc, Transform linkToHost, string linkName)
        {
            return new LinkedRoomItem { Room = room, SourceDoc = linkDoc, LinkToHost = linkToHost ?? Transform.Identity, IsFromLink = true, LinkInstanceName = linkName };
        }

        public override string ToString() => DisplayText;
    }
}