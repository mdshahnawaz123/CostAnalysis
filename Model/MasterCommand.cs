using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CostAnalysis.Model
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MasterCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uidoc = commandData.Application.ActiveUIDocument;
            var doc = uidoc.Document;

            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .WhereElementIsNotElementType()
                .ToList();

            foreach ( var element in collector )
            {
                var family = element as Family;
                if ( family != null )
                {
                    var nam = family.Name;
                }
            }
            TaskDialog.Show("Message","");
            return Result.Succeeded;
        }
    }
}
