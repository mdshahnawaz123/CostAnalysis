using System;
using System.Windows.Interop;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using CostAnalysis.UI;

namespace CostAnalysis.Command
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PramCheck : Autodesk.Revit.UI.IExternalCommand
    {
        public static DataExporter Instance;

        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            var uidoc = commandData.Application.ActiveUIDocument;
            var doc = uidoc.Document;

            // ── LOGIN CHECK ───────────────────────────────────────────────────
            if (!RevitUI.UI.LoginGuard.IsAuthorized()) return Autodesk.Revit.UI.Result.Cancelled;

            try
            {
                var revitHandle = GetRevitHandle(commandData);
                return RunMainWindow(doc, uidoc, revitHandle);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Autodesk.Revit.UI.Result.Failed;
            }
        }

        private Autodesk.Revit.UI.Result RunMainWindow(Autodesk.Revit.DB.Document doc, Autodesk.Revit.UI.UIDocument uidoc, IntPtr revitHandle)
        {
            if (Instance != null)
            {
                Instance.Activate();
                if (Instance.WindowState == System.Windows.WindowState.Minimized)
                    Instance.WindowState = System.Windows.WindowState.Normal;
                return Autodesk.Revit.UI.Result.Succeeded;
            }

            Instance = new DataExporter(doc, uidoc);
            Instance.Closed += (s, e) => { Instance = null; };
            Instance.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
            new WindowInteropHelper(Instance) { Owner = revitHandle };
            Instance.Show();

            return Autodesk.Revit.UI.Result.Succeeded;
        }

        private IntPtr GetRevitHandle(ExternalCommandData commandData)
        {
            try
            {
                var handle = new IntPtr((int)commandData.Application.MainWindowHandle);
                if (handle != IntPtr.Zero) return handle;
            }
            catch { /* ignore */ }
            return System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
        }
    }
}
