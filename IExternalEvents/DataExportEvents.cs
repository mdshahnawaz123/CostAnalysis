using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections;

public class DataExportEvents : IExternalEventHandler
{
    
    public void Execute(UIApplication app)
    {
        try
        {
            TaskDialog.Show("Handler", "Execute() started");

            var uidoc = app?.ActiveUIDocument;
            if (uidoc == null)
            {
                TaskDialog.Show("Handler", "ActiveUIDocument is null");
                return;
            }

            var doc = uidoc.Document;


        }
        catch (Exception ex)
        {
            // show the exception so you don't miss it
            TaskDialog.Show("Handler - Exception", ex.ToString());
        }
    }

    public string GetName() => "Data Export Event Handler";
}
