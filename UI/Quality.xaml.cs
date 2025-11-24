using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace CostAnalysis.UI
{
    public partial class Quality : Window
    {
        private readonly Document _doc;
        private readonly UIDocument _uidoc;

        public Quality(Document doc, UIDocument uidoc)
        {
            InitializeComponent();
            _doc = doc;
            _uidoc = uidoc;
            info();
        }

        // Load & bind tree
        public void info()
        {
            Tv_Explore.Items.Clear();

            // Expect top-level categories from your helper
            IEnumerable<Category> revitCategories = Extension.DataLab.Categories(_doc);

            // Map to view-model nodes (with counts & BuiltInCategory)
            var nodes = revitCategories
                .Select(c => MapCategoryToNode(_doc, c))
                .OrderBy(n => n.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            Tv_Explore.ItemsSource = nodes;
        }

        // ========== Mapping ==========
        private static CategoryNode MapCategoryToNode(Document doc, Category cat)
        {
            var node = new CategoryNode
            {
                Name = cat?.Name ?? "(Unnamed)",
                Count = SafeCountElements(doc, cat),
                Bic = TryGetBic(cat)
            };

            if (cat?.SubCategories != null && cat.SubCategories.Size > 0)
            {
                foreach (Category sub in cat.SubCategories)
                    node.Children.Add(MapCategoryToNode(doc, sub));
            }
            return node;
        }

        private static BuiltInCategory? TryGetBic(Category cat)
        {
            try { return (BuiltInCategory)cat?.Id.IntegerValue; }
            catch { return null; }
        }

        private static int SafeCountElements(Document doc, Category cat)
        {
            try
            {
                var bic = (BuiltInCategory)cat?.Id.IntegerValue;
                return new FilteredElementCollector(doc)
                        .OfCategory(bic)
                        .WhereElementIsNotElementType()
                        .ToElementIds()
                        .Count;
            }
            catch { return 0; }
        }

        // ========== Highlight ==========
        private void Btn_HighLight_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_uidoc == null || _doc == null) return;

                var roots = Tv_Explore.ItemsSource as IEnumerable<CategoryNode>;
                if (roots == null) return;

                // gather checked nodes
                var checkedNodes = roots.SelectMany(GetCheckedNodesRecursive).ToList();
                if (!checkedNodes.Any())
                {
                    TaskDialog.Show("Parameter QC", "No categories are checked.");
                    return;
                }

                // collect element ids
                var ids = new HashSet<ElementId>();
                foreach (var n in checkedNodes)
                    foreach (var id in CollectIdsForNode(_doc, n))
                        ids.Add(id);

                if (ids.Count == 0)
                {
                    TaskDialog.Show("Parameter QC", "No elements found for the selected categories.");
                    return;
                }

                // select + zoom
                _uidoc.Selection.SetElementIds(ids);
                _uidoc.ShowElements(ids);

                // optionally: temporary isolate
                using (var t = new Transaction(_doc, "Temporary Isolate"))
                {
                    t.Start();
                    _doc.ActiveView.IsolateElementsTemporary(ids);
                    t.Commit();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Parameter QC", ex.Message);
            }
        }

        private static IEnumerable<CategoryNode> GetCheckedNodesRecursive(CategoryNode node)
        {
            if (node.IsChecked == true)
                yield return node;

            if (node.Children == null) yield break;

            foreach (var c in node.Children.SelectMany(GetCheckedNodesRecursive))
                yield return c;
        }

        private static IEnumerable<CategoryNode> GetCheckedNodesRecursive(IEnumerable<CategoryNode> roots)
        {
            foreach (var r in roots)
                foreach (var n in GetCheckedNodesRecursive(r))
                    yield return n;
        }

        private static IEnumerable<ElementId> CollectIdsForNode(Document doc, CategoryNode node)
        {
            if (node.Bic == null) yield break;

            var collector = new FilteredElementCollector(doc)
                .OfCategory(node.Bic.Value)
                .WhereElementIsNotElementType();

            foreach (var e in collector)
                yield return e.Id;
        }
    }

    // ===== View-model used by TreeView =====
    public class CategoryNode
    {
        public string Name { get; set; }
        public List<CategoryNode> Children { get; set; } = new List<CategoryNode>();
        public bool? IsChecked { get; set; } = false;

        // shown as (Count)
        public int Count { get; set; }

        // reference to Revit category for fast collection
        public BuiltInCategory? Bic { get; set; }
    }
}
