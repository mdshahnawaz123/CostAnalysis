using Autodesk.Revit.UI;
using CostAnalysis.Services;
using System;
using System.IO;
using System.Reflection;
using System.Windows;

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
            var source = "https://raw.githubusercontent.com/mdshahnawaz123/plugin-access-control/main/users.json";

            try
            {
                var auth = new AuthService(source);
                CostAnalysis.Command.PramCheck.Auth = auth;
            }
            catch
            {
                // ... log error
            }

            try
            {
                RibbonPanel panel = FindExistingPanel(application, PanelName);
                if (panel == null)
                {
                    if (!TabExists(application, TargetTabName))
                    {
                        try { application.CreateRibbonTab(TargetTabName); }
                        catch { }
                    }
                    panel = application.CreateRibbonPanel(TargetTabName, PanelName);
                }

                // Add Data Exporter Button
                AddPushButton(panel, "QC_Analysis_DataExporter", "Data Exporter", "CostAnalysis.Command.PramCheck", "Open Data Exporter (BOQ & QA-QC)", "Help.html");

                // Add Quality Check Button
                AddPushButton(panel, "QC_Analysis_Quality", "Quality Check", "CostAnalysis.Command.QualityCommand", "Open Parameter Quality Check tool", "Help.html");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // ... log error
                ex.Message.ToString();
                return Result.Succeeded;
            }
        }

        private void AddPushButton(RibbonPanel panel, string internalName, string text, string className, string tooltip, string helpFileName = null)
        {
            if (PanelHasButton(panel, internalName)) return;

            var pushData = new PushButtonData(internalName, text, AssemblyPath, className)
            {
                ToolTip = tooltip
            };

            if (!string.IsNullOrEmpty(helpFileName))
            {
                var helpPath = Path.Combine(Path.GetDirectoryName(AssemblyPath), "Resources", helpFileName);
                if (File.Exists(helpPath))
                {
                    var help = new ContextualHelp(ContextualHelpType.ChmFile, helpPath);
                    pushData.SetContextualHelp(help);
                }
            }

            var item = panel.AddItem(pushData) as PushButton;
            if (item != null)
            {
                try
                {
                    item.LargeImage = LoadImageFromResource("Resources/Icon32.png");
                    item.Image = LoadImageFromResource("Resources/Icon16.png");
                }
                catch { }
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
