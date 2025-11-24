using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CostAnalysis.Extension;
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
            var app = doc.Application;

            //Lets Talk about Shared Parameters:
            var definactionFile = app.OpenSharedParameterFile(); //Lets check sharedfile are there are not:

            if(definactionFile == null)
            {
                TaskDialog.Show("Error", "No Shared Parameter File is assigned in Revit Options, Please Create the Shared-Parameter ");
                return Result.Failed;
            }

            //Lets Check for the Group:

            var group = definactionFile.Groups.get_Item("ECD_ABS_Parameters");

            //Lets Create the Shared Parameters if not exist:




            return Result.Succeeded;
        }
    }
}
