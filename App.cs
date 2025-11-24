using Autodesk.Revit.UI;
using CostAnalysis.Services;
using System;
using System.IO;
using System.Reflection;

namespace CostAnalysis
{
    public class App : IExternalApplication
    {
        private const string TargetTabName = "BIM Digital Design";
        private const string PanelName = "QC Panel";
        private const string ButtonInternalName = "QC_Analysis_DataExporter";
        private const string ButtonText = "Data Exporter";
        private const string ButtonTooltip = "Open Data Exporter (BOQ & QA-QC)";
        private const string commandClass = "CostAnalysis.Command.PramCheck";

        private string AssemblyPath => Assembly.GetExecutingAssembly().Location;

        public Result OnStartup(UIControlledApplication application)
        {
            // Use uploaded path (will be transformed if you replace it with the raw URL)


            // production example:
            var source = "https://raw.githubusercontent.com/mdshahnawaz123/plugin-access-control/main/users.json";

            try
            {
                var auth = new AuthService(source);
                CostAnalysis.Command.PramCheck.Auth = auth;
            }
            catch
            {
                try
                {
                    var folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CostAnalysis");
                    Directory.CreateDirectory(folder);
                    var file = Path.Combine(folder, "startup_auth.log");
                    File.AppendAllText(file, DateTime.UtcNow.ToString("s") + " AuthService init failed." + Environment.NewLine);
                }
                catch { }
            }

            try
            {
                RibbonPanel panel = FindExistingPanel(application, PanelName);
                if (panel == null)
                {
                    try
                    {
                        if (!TabExists(application, TargetTabName))
                        {
                            try { application.CreateRibbonTab(TargetTabName); }
                            catch { }
                        }
                        panel = application.CreateRibbonPanel(TargetTabName, PanelName);
                    }
                    catch
                    {
                        panel = application.CreateRibbonPanel(PanelName);
                    }
                }

                if (!PanelHasButton(panel, ButtonInternalName))
                {
                    var pushData = new PushButtonData(ButtonInternalName, ButtonText, AssemblyPath, commandClass)
                    {
                        ToolTip = ButtonTooltip
                    };

                    var item = panel.AddItem(pushData);
                    var push = item as PushButton;

                    try
                    {
                        var large = LoadImageFromResource("Resources/Icon32.png");
                        if (large != null && push != null) push.LargeImage = large;
                        var small = LoadImageFromResource("Resources/Icon16.png");
                        if (small != null && push != null) push.Image = small;
                    }
                    catch { }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                try
                {
                    var folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CostAnalysis");
                    Directory.CreateDirectory(folder);
                    var file = Path.Combine(folder, "startup.log");
                    File.AppendAllText(file, DateTime.UtcNow.ToString("s") + " OnStartup exception: " + ex + Environment.NewLine);
                }
                catch { }
                return Result.Succeeded;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        // Helper methods below (same as you used before)
        private bool TabExists(UIControlledApplication app, string tabName)
        {
            try
            {
                var panels = app.GetRibbonPanels(tabName);
                return panels != null && panels.Count > 0;
            }
            catch { return false; }
        }

        private RibbonPanel FindExistingPanel(UIControlledApplication app, string panelName)
        {
            try
            {
                var allPanels = app.GetRibbonPanels();
                foreach (var p in allPanels)
                {
                    if (string.Equals(p.Name, panelName, StringComparison.OrdinalIgnoreCase))
                        return p;
                }
            }
            catch { }
            return null;
        }

        private bool PanelHasButton(RibbonPanel panel, string internalName)
        {
            if (panel == null) return false;
            try
            {
                var items = panel.GetItems();
                foreach (var it in items)
                {
                    if (string.Equals(it.Name, internalName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch { }
            return false;
        }

        private System.Windows.Media.ImageSource LoadImageFromResource(string relativeResourcePath)
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var asmName = asm.GetName().Name;
                var packUri = $"pack://application:,,,/{asmName};component/{relativeResourcePath}";
                var bmp = new System.Windows.Media.Imaging.BitmapImage();
                bmp.BeginInit();
                bmp.UriSource = new Uri(packUri, UriKind.Absolute);
                bmp.EndInit();
                bmp.Freeze();
                return bmp;
            }
            catch
            {
                try
                {
                    var asm = Assembly.GetExecutingAssembly();
                    var names = asm.GetManifestResourceNames();
                    string match = Array.Find(names, n => n.EndsWith(relativeResourcePath.Replace('/', '.').Replace('\\', '.'), StringComparison.OrdinalIgnoreCase));
                    if (string.IsNullOrEmpty(match)) return null;
                    using (var s = asm.GetManifestResourceStream(match))
                    {
                        if (s == null) return null;
                        var bmp = new System.Windows.Media.Imaging.BitmapImage();
                        bmp.BeginInit();
                        bmp.StreamSource = s;
                        bmp.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
                        bmp.EndInit();
                        bmp.Freeze();
                        return bmp;
                    }
                }
                catch { return null; }
            }
        }
    }
}
