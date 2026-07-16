using Autodesk.Revit.UI;

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

        private string AssemblyPath => Assembly.GetExecutingAssembly().Location;

        public Result OnStartup(UIControlledApplication application)
        {
            // ── Ribbon ──────────────────────────────────────────────────
            try
            {
                // Create tab only if it doesn't exist yet
                if (!TabExists(application, TargetTabName))
                {
                    try { application.CreateRibbonTab(TargetTabName); }
                    catch { }
                }

                // Find existing panel or create a new one
                RibbonPanel panel = FindExistingPanel(application, PanelName)
                                    ?? application.CreateRibbonPanel(TargetTabName, PanelName);

                // Buttons
                AddPushButton(panel,
                    "QC_Analysis_DataExporter",
                    "Data Exporter",
                    "CostAnalysis.Command.PramCheck",
                    "Open Data Exporter (BOQ & QA-QC)",
                    "Help.html");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                File.WriteAllText(@"C:\Temp\CostAnalysis_error.txt", ex.ToString());
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        // ── Tab helpers ─────────────────────────────────────────────────

        /// <summary>
        /// Returns true if the tab already exists.
        /// Relies on GetRibbonPanels() throwing when the tab is unknown.
        /// </summary>
        private bool TabExists(UIControlledApplication app, string tabName)
        {
            try
            {
                app.GetRibbonPanels(tabName); // throws if tab doesn't exist
                return true;
            }
            catch { return false; }
        }

        /// <summary>
        /// Searches the target tab for a panel with the given name.
        /// Returns null if the tab or panel doesn't exist yet.
        /// </summary>
        private RibbonPanel FindExistingPanel(UIControlledApplication app, string panelName)
        {
            try
            {
                if (!TabExists(app, TargetTabName)) return null;

                foreach (var p in app.GetRibbonPanels(TargetTabName))
                {
                    if (string.Equals(p.Name, panelName, StringComparison.OrdinalIgnoreCase))
                        return p;
                }
            }
            catch { }
            return null;
        }

        // ── Button helpers ──────────────────────────────────────────────

        private void AddPushButton(RibbonPanel panel, string internalName, string text,
                                   string className, string tooltip, string helpFileName = null)
        {
            if (PanelHasButton(panel, internalName)) return;

            var pushData = new PushButtonData(internalName, text, AssemblyPath, className)
            {
                ToolTip = tooltip
            };

            if (!string.IsNullOrEmpty(helpFileName))
            {
                var helpPath = Path.Combine(
                    Path.GetDirectoryName(AssemblyPath), "Resources", helpFileName);
                if (File.Exists(helpPath))
                    pushData.SetContextualHelp(
                        new ContextualHelp(ContextualHelpType.ChmFile, helpPath));
            }

            var item = panel.AddItem(pushData) as PushButton;
            if (item != null)
            {
                try { item.LargeImage = LoadImageFromResource("Resources/Icon32.png"); } catch { }
                try { item.Image = LoadImageFromResource("Resources/Icon16.png"); } catch { }
            }
        }

        private bool PanelHasButton(RibbonPanel panel, string internalName)
        {
            if (panel == null) return false;
            try
            {
                foreach (var it in panel.GetItems())
                    if (string.Equals(it.Name, internalName, StringComparison.OrdinalIgnoreCase))
                        return true;
            }
            catch { }
            return false;
        }

        // ── Image loader ────────────────────────────────────────────────

        private System.Windows.Media.ImageSource LoadImageFromResource(string relativeResourcePath)
        {
            // 1. Try pack URI (works when resources are marked as Resource in .csproj)
            try
            {
                var asmName = Assembly.GetExecutingAssembly().GetName().Name;
                var uri = new Uri($"pack://application:,,,/{asmName};component/{relativeResourcePath}",
                                      UriKind.Absolute);
                var bmp = new System.Windows.Media.Imaging.BitmapImage();
                bmp.BeginInit();
                bmp.UriSource = uri;
                bmp.EndInit();
                bmp.Freeze();
                return bmp;
            }
            catch { }

            // 2. Fallback: embedded manifest resource stream
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var key = relativeResourcePath.Replace('/', '.').Replace('\\', '.');
                var match = Array.Find(asm.GetManifestResourceNames(),
                                       n => n.EndsWith(key, StringComparison.OrdinalIgnoreCase));
                if (match == null) return null;

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