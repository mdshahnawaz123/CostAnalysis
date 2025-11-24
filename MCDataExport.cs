using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CostAnalysis
{
    [Transaction(TransactionMode.Manual)]
    public class MCDataExport : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uidoc = commandData.Application.ActiveUIDocument;
            var doc = uidoc.Document;

            var eleid = uidoc.Selection.GetElementIds();

           foreach ( var ele in eleid )
            {
                var eleName = doc.GetElement(ele);


                foreach (Parameter p in eleName.Parameters)
                {
                    string name = p.Definition.Name;
                    string value = p.HasValue ? p.AsValueString() : "No Value";
                    
                }



                TaskDialog.Show("Message", eleName.ToString());

            }



            return Result.Succeeded;
        }
    }
}
