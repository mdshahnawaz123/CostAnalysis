using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CostAnalysis.Extension
{
    public static class DataLab
    {
        public static IList<Category> categories(Document doc)
        {
            var cat = new List<Category>();

            var categor = doc.Settings.Categories;

            foreach (Category c in categor)
            {
                cat.Add(c);
            }
            return cat.OrderBy(x => x.Name).ToList(); ;
        }

        public static void ShowElement(UIDocument uidoc,ICollection<Element> elements)
        {
            if (elements == null || elements.Count == 0)
            {
                TaskDialog.Show("Error", "Please select elements to highlight");
                return;
            }

            ICollection<ElementId> ids = new List<ElementId>();

            foreach (var ele in elements)
            {
                ids.Add(ele.Id);
            }
            uidoc.ShowElements(ids);
        }

    }
}
