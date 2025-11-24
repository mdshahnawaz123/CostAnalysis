using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CostAnalysis.Extension
{
    public static class DataLab
    {
        // Canonical list of ABS parameter names (single source of truth)
        public static readonly string[] AbsParamNames = new[]
        {
            "ECD_ABS_L1_Asset",
            "ECD_ABS_L1_Contract",
            "ECD_ABS_L1_WBS_Plot_Code",
            "ECD_ABS_L2_Level",
            "ECD_ABS_L3_Room",
            "ECD_ABS_L4_Equipment_Group",
            "ECD_ABS_L5_Equipment_System",
            "ECD_ABS_L5_OM_Manual_Ref",
            "ECD_ABS_L6_Equipment_Type",
            "ECD_ABS_L6_Manufacturer_Name",
            "ECD_ABS_L6_Supplier_Name",
            "ECD_ABS_L6_Make",
            "ECD_ABS_L6_Model",
            "ECD_ABS_L6_Life_Expectancy",
            "ECD_ABS_L6_Quantity",
            "ECD_ABS_L7_Equipment_Unique_Number",
            "ECD_ABS_L7_Onsite_Equipment_Tag"
        };

        public static IList<Category> Categories(Document doc)
        {
            var cat = new List<Category>();
            var categor = doc.Settings.Categories;
            foreach (Category c in categor) cat.Add(c);
            return cat.OrderBy(x => x.Name).ToList();
        }

        public static void ShowElement(UIDocument uidoc, ICollection<Element> elements)
        {
            if (elements == null || elements.Count == 0)
            {
                TaskDialog.Show("Error", "Please select elements to highlight");
                return;
            }

            ICollection<ElementId> ids = new List<ElementId>();
            foreach (var ele in elements) ids.Add(ele.Id);
            uidoc.ShowElements(ids);
        }

        public static IList<Element> ElementByCat(Document doc, Category cat)
        {
            

            var ele = new FilteredElementCollector(doc)
                .OfCategoryId(cat.Id)
                .WhereElementIsNotElementType()
                .ToList();
            return ele;
        }
        public static IList<Category> GetUsedCategories(Document doc)
        {
            // Collect all placed (instance) elements
            var instanceCollector = new FilteredElementCollector(doc)
                                        .WhereElementIsNotElementType();

            // Extract categories from those elements
            var usedCategories = instanceCollector
                                    .Select(e => e.Category)
                                    .Where(c => c != null)
                                    .Where(c => c.CategoryType == CategoryType.Model)  // only model categories
                                    .GroupBy(c => c.Id.IntegerValue)
                                    .Select(g => g.First())
                                    .OrderBy(c => c.Name)
                                    .ToList();

            return usedCategories;
        }

        public static SharedParameterElement GetSharedParameterElement(Document doc, string paramName)
        {
            try
            {
                var sharedpath = doc.Application.SharedParametersFilename;
                if (string.IsNullOrEmpty(sharedpath))
                    return null;

                var collector = new FilteredElementCollector(doc);
                var element = collector
                    .OfClass(typeof(SharedParameterElement))
                    .Cast<SharedParameterElement>()
                    .FirstOrDefault(p => p.Name == paramName);

                return element;
            }
            catch
            {
                return null;
            }
        }


        public class ParamInfo
        {
            public string TypeMark { get; set; }
            public double length { get; set; }
            public double width { get; set; }
            public double height { get; set; }
            public double Area { get; set; }
            public double Volume { get; set; }
            public string Reinfo_Data { get; set; }
            public string opening_Data { get; set; }
            public string functional_Use { get; set; }
            public Material material { get; set; }
            public Workset workset { get; set; }
        }
    }
}