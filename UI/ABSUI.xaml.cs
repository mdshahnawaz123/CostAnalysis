using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace CostAnalysis.UI
{
    public partial class ABSUI : Window
    {
        UIDocument uidoc;
        Document doc;

        // We’ll list both host and linked rooms via this wrapper
        private List<LinkedRoomItem> _allRooms = new List<LinkedRoomItem>();

        public ABSUI(Document doc, UIDocument uidoc)
        {
            InitializeComponent();
            this.doc = doc ?? throw new ArgumentNullException(nameof(doc));
            this.uidoc = uidoc ?? throw new ArgumentNullException(nameof(uidoc));

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
                        "Could not compute a 3D solid for the selected room.\n" +
                        "Ensure the room is properly enclosed in its phase.");
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

        public void Btn_ABS(object sender, RoutedEventArgs e)
        {
            try
            {
                // TODO: assign your ABS code to selected elements
                // var code = AssetCodeText.Text;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("ABS", $"ABS failed.\n\n{ex.Message}");
            }
        }

        public void Btn_Export(object sender, RoutedEventArgs e)
        {
            try
            {
                // TODO: export your data
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

            foreach (var r in hostRooms)
            {
                items.Add(LinkedRoomItem.FromHost(r, doc));
            }

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
                {
                    items.Add(LinkedRoomItem.FromLink(r, linkDoc, linkToHost, linkInst.Name));
                }
            }

            _allRooms = items
                .OrderBy(i => i.IsFromLink)            // host first
                .ThenBy(i => i.SourceTag)              // group by source
                .ThenBy(i => i.Number)
                .ThenBy(i => i.Name)
                .ToList();

            RoomsList.ItemsSource = _allRooms;

            // Show a friendly label in the ListBox
            RoomsList.DisplayMemberPath = nameof(LinkedRoomItem.DisplayText);
        }

        private void BuildCategoryTree()
        {
            // Fill if/when you hook up the right-side tree
        }

        // ------ Geometry -----------------------------------------------------------

        /// <summary>
        /// Build a 3D solid from the room's boundary loops by extruding upward.
        /// Works for both host and linked rooms (we transform later for linked).
        /// </summary>
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

            // Height: prefer UnboundedHeight; fallback to ~3m in feet
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
            catch
            {
                // ignore
            }

            return null;
        }
    }

    /// <summary>
    /// Wrapper to represent a room that may live in host or a linked document.
    /// Provides the transform from the room's doc into the host doc.
    /// </summary>
    internal class LinkedRoomItem
    {
        public Room Room { get; private set; }
        public Document SourceDoc { get; private set; }
        public Transform LinkToHost { get; private set; } // Identity for host
        public bool IsFromLink { get; private set; }
        public string LinkInstanceName { get; private set; }

        public string Number => Room?.Number;
        public string Name => Room?.Name;

        public string SourceTag => IsFromLink
            ? (string.IsNullOrEmpty(LinkInstanceName) ? "Link" : LinkInstanceName)
            : "Host";

        public string DisplayText
            => $"{Number} - {Name}  [{SourceTag}]";

        private LinkedRoomItem() { }

        public static LinkedRoomItem FromHost(Room room, Document hostDoc)
        {
            return new LinkedRoomItem
            {
                Room = room,
                SourceDoc = hostDoc,
                LinkToHost = Transform.Identity,
                IsFromLink = false,
                LinkInstanceName = "Host"
            };
        }

        public static LinkedRoomItem FromLink(Room room, Document linkDoc, Transform linkToHost, string linkName)
        {
            return new LinkedRoomItem
            {
                Room = room,
                SourceDoc = linkDoc,
                LinkToHost = linkToHost ?? Transform.Identity,
                IsFromLink = true,
                LinkInstanceName = linkName
            };
        }

        public override string ToString() => DisplayText;
    }
}
