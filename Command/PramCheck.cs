using System;
using System.Windows.Interop;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using CostAnalysis.Services;
using CostAnalysis.UI;

namespace CostAnalysis.Command
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PramCheck : Autodesk.Revit.UI.IExternalCommand
    {
        public static AuthService Auth { get; set; }

        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            var uidoc = commandData.Application.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
                if (Auth == null)
                {
                    message = "Authentication service not initialized.";
                    return Autodesk.Revit.UI.Result.Failed;
                }

                var loaded = Auth.TryLoadRemote().GetAwaiter().GetResult();
                if (!loaded)
                {
                    TaskDialog.Show("Network required", "This addin requires internet to validate users. Please connect and try again.");
                    return Autodesk.Revit.UI.Result.Cancelled;
                }

                var token = TokenService.LoadToken();
                var machineId = MachineHelper.GetMachineId();

                if (token != null && token.MachineId == machineId && token.ExpiresUtc.Date >= DateTime.UtcNow.Date)
                {
                    var remoteUser = Auth.GetUser(token.Username);
                    if (remoteUser != null && remoteUser.Active && remoteUser.Expires.Date >= DateTime.UtcNow.Date)
                    {
                        Auth.CurrentUser = remoteUser;
                        var revitHandle = GetRevitHandle(commandData);
                        return RunMainWindow(doc, uidoc, revitHandle);
                    }
                    else
                    {
                        TokenService.DeleteToken();
                    }
                }

                var login = new LoginWindow(Auth);
                var revitHandleLogin = GetRevitHandle(commandData);
                new WindowInteropHelper(login) { Owner = revitHandleLogin };

                var dlg = login.ShowDialog();
                if (dlg != true)
                {
                    return Autodesk.Revit.UI.Result.Cancelled;
                }

                var user = Auth.CurrentUser;
                if (user == null)
                {
                    TaskDialog.Show("Login error", "Failed to retrieve user after login.");
                    return Autodesk.Revit.UI.Result.Failed;
                }

                var localToken = new LocalAuthToken
                {
                    Username = user.Username,
                    MachineId = machineId,
                    ExpiresUtc = user.Expires.ToUniversalTime()
                };
                TokenService.SaveToken(localToken);

                var revitHandleAfterLogin = GetRevitHandle(commandData);
                return RunMainWindow(doc, uidoc, revitHandleAfterLogin);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Autodesk.Revit.UI.Result.Failed;
            }
        }

        private Autodesk.Revit.UI.Result RunMainWindow(Autodesk.Revit.DB.Document doc, Autodesk.Revit.UI.UIDocument uidoc, IntPtr revitHandle)
        {
            var user = Auth.CurrentUser;
            if (!IsUserAllowed(user))
            {
                TaskDialog.Show("Access denied", "You do not have permission to run this tool.");
                return Autodesk.Revit.UI.Result.Failed;
            }

            var frm = new DataExporter(doc, uidoc);
            frm.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
            new WindowInteropHelper(frm) { Owner = revitHandle };
            frm.Show();

            return Autodesk.Revit.UI.Result.Succeeded;
        }

        private bool IsUserAllowed(Model.UserRecord user)
        {
            if (user == null) return false;
            return user.Active && user.Expires.Date >= DateTime.UtcNow.Date;
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
