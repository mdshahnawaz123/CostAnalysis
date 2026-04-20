// users.json dev reference (uploaded file path):
// /mnt/data/b6955248-e5d2-4f87-bf5b-499886feb251.png
//
// Paste this file as UI/DataExporter.xaml.cs

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Runtime.InteropServices;
using System.Windows.Navigation;

namespace CostAnalysis.UI
{
    public partial class DataExporter : Window, INotifyPropertyChanged
    {
        private ExternalEvent _externalEvent;
        private CostAnalysis.IExternalEvents.DataExportEvents _eventHandler;

        public ObservableCollection<TypeEntry> TypeEntries { get; set; } = new ObservableCollection<TypeEntry>();
        // verbose info (not bound to the LB for export) - useful for debugging or tooltip
        public ObservableCollection<string> ParameterInfo { get; set; } = new ObservableCollection<string>();
        // raw parameter names used by the QA-QC ListBox (LB_TypeParameters)
        public ObservableCollection<string> ParameterNames { get; set; } = new ObservableCollection<string>();

        public ObservableCollection<string> GeometryPreview { get; set; } = new ObservableCollection<string>();

        private ICollectionView _typesView;
        private string _typeSearchText = "";

        Document doc;
        UIDocument uidoc;

        private TypeEntry _selectedType;
        public TypeEntry SelectedType
        {
            get => _selectedType;
            set
            {
                if (_selectedType == value) return;
                _selectedType = value;
                OnPropertyChanged(nameof(SelectedType));
                UpdateSelectedTypeDetails();
            }
        }

        public DataExporter(Document doc, UIDocument uidoc)
        {
            InitializeComponent();
            this.doc = doc;
            this.uidoc = uidoc;

            DataContext = this;

            // Initialize external event for context-safe modifications
            _eventHandler = new CostAnalysis.IExternalEvents.DataExportEvents();
            _eventHandler.ProgressUpdater = UpdateProgress;
            _externalEvent = ExternalEvent.Create(_eventHandler);

            // use an ICollectionView for stable filtering
            _typesView = CollectionViewSource.GetDefaultView(TypeEntries);
            _typesView.Filter = TypesFilter;

            DG_Types.ItemsSource = _typesView;
            LB_TypeParameters.ItemsSource = ParameterNames; // bind QA-QC list to raw names
            LV_GeometryPreview.ItemsSource = GeometryPreview;

            // wire double-click toggle on parameter list (so user can double-click to toggle selection)
            LB_TypeParameters.MouseDoubleClick += LB_TypeParameters_MouseDoubleClick;

            // populate categories (DataLab helper)
            try
            {
                LV_ModelCat.ItemsSource = CostAnalysis.Extension.DataLab.GetUsedCategories(doc);
            }
            catch
            {
                LV_ModelCat.ItemsSource = null;
            }

            LV_ModelCat.SelectionChanged += LV_ModelCat_SelectionChanged;
            DG_Types.SelectionChanged += DG_Types_SelectionChanged;
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            IconHelper.RemoveIcon(this);
        }

        #region INotify
        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string n) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
        #endregion

        #region UI Handlers

        public void UpdateProgress(int current, int total, string message)
        {
            if (current < total && total > 0)
            {
                PB_Progress.Visibility = System.Windows.Visibility.Visible;
                TXT_Progress.Visibility = System.Windows.Visibility.Visible;
                PB_Progress.Maximum = total;
                PB_Progress.Value = current;
                TXT_Progress.Text = message;
            }
            else
            {
                PB_Progress.Visibility = System.Windows.Visibility.Hidden;
                TXT_Progress.Visibility = System.Windows.Visibility.Hidden;
            }

            // Force layout and render updates
            System.Windows.Application.Current.Dispatcher.Invoke(
                new Action(() => { }),
                System.Windows.Threading.DispatcherPriority.Background);
        }

        private void LV_ModelCat_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var selectedCates = LV_ModelCat.SelectedItems.Cast<Category>().ToList();
            if (!selectedCates.Any())
            {
                TypeEntries.Clear();
                ParameterNames.Clear();
                ParameterInfo.Clear();
                GeometryPreview.Clear();
                return;
            }
            BuildTypesForCategories(selectedCates);
        }

        private void DG_Types_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            SelectedType = DG_Types.SelectedItem as TypeEntry;
        }

        private bool TypesFilter(object obj)
        {
            if (obj is TypeEntry te)
            {
                if (string.IsNullOrEmpty(_typeSearchText)) return true;
                return (te.Name ?? "").IndexOf(_typeSearchText, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            return false;
        }

        private void TB_TypeSearch_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            _typeSearchText = TB_TypeSearch.Text?.Trim() ?? "";
            _typesView.Refresh();

            if (SelectedType != null && !_typesView.Cast<TypeEntry>().Contains(SelectedType))
            {
                DG_Types.SelectedItem = null;
                SelectedType = null;
            }
        }

        private void CB_SelectAllTypes_Checked(object sender, RoutedEventArgs e)
        {
            foreach (var t in TypeEntries) t.IsSelected = true;
            DG_Types.Items.Refresh();
        }

        private void CB_SelectAllTypes_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (var t in TypeEntries) t.IsSelected = false;
            DG_Types.Items.Refresh();
        }

        public void Close_Click(object sender, RoutedEventArgs e) => Close();

        public void HighLight_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selcats = LV_ModelCat.SelectedItems.Cast<Category>().ToList();
                if (!selcats.Any())
                {
                    MessageBox.Show("Select at least one model category first.", "No Category", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var ids = new List<ElementId>();
                foreach (var cat in selcats)
                {
                    ids.AddRange(new FilteredElementCollector(doc).OfCategoryId(cat.Id).WhereElementIsNotElementType().ToElementIds());
                }

                if (!ids.Any()) return;

                uidoc.Selection.SetElementIds(ids);
                uidoc.ShowElements(ids);
                if (doc.ActiveView.ViewType == ViewType.ThreeD || doc.ActiveView.ViewType == ViewType.FloorPlan)
                {
                    using (Transaction t = new Transaction(doc, "Temp Isolate"))
                    {
                        t.Start();
                        doc.ActiveView.IsolateElementsTemporary(ids);
                        t.Commit();
                    }
                }
                else MessageBox.Show("Isolation works only in 3D or plan views.");
            }
            catch (Exception ex) { MessageBox.Show("Highlight failed: " + ex.Message + "\n\nPlease contact support: mohammad.shahnawaz@expocitydubai.ae"); }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F1)
            {
                try
                {
                    var helpPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Resources", "Help.html");
                    if (System.IO.File.Exists(helpPath))
                    {
                        Process.Start(new ProcessStartInfo(helpPath) { UseShellExecute = true });
                    }
                }
                catch
                {
                    // Ignore exceptions if help fails to open
                }
                e.Handled = true;
            }
        }

        #endregion

        #region Build lists

        private void BuildTypesForCategories(List<Category> selCats)
        {
            TypeEntries.Clear();
            ParameterNames.Clear();
            ParameterInfo.Clear();
            GeometryPreview.Clear();

            var combinedElems = new List<Element>();
            foreach (var cat in selCats)
            {
                var elems = new FilteredElementCollector(doc)
                    .OfCategoryId(cat.Id)
                    .WhereElementIsNotElementType()
                    .Where(e => e != null)
                    .ToList();
                combinedElems.AddRange(elems);
            }

            var groups = combinedElems.GroupBy(x => x.GetTypeId())
                              .Select(g =>
                              {
                                  var typeId = g.Key;
                                  var typeElem = doc.GetElement(typeId);
                                  string name = typeElem != null ? (typeElem.Name ?? typeElem.GetType().Name) : $"Type {typeId.Value}";

                                  var instances = new ObservableCollection<InstanceEntry>();
                                  foreach (var el in g)
                                  {
                                      instances.Add(new InstanceEntry
                                      {
                                          Id = el.Id.Value.ToString(),
                                          Name = TryGetElementName(el),
                                          Level = TryGetLevelName(el),
                                          InfoSummary = BuildInstanceInfoSummary(el),
                                          ElementRef = el
                                      });
                                  }

                                  return new TypeEntry
                                  {
                                      Name = name,
                                      TypeId = typeId,
                                      RepresentativeTypeElement = typeElem,
                                      Instances = instances
                                  };
                              })
                              .OrderBy(t => t.Name);

            foreach (var te in groups) TypeEntries.Add(te);

            _typesView.Refresh();

            if (TypeEntries.Any())
            {
                DG_Types.SelectedItem = TypeEntries.First();
                SelectedType = TypeEntries.First();
            }
        }

        /// <summary>
        /// Populate ParameterNames (raw names) and ParameterInfo (verbose) using both Type and Instance parameters.
        /// </summary>
        private void UpdateSelectedTypeDetails()
        {
            ParameterNames.Clear();
            ParameterInfo.Clear();
            GeometryPreview.Clear();

            if (SelectedType == null) return;

            var repType = SelectedType.RepresentativeTypeElement;
            var repInst = SelectedType.Instances.FirstOrDefault()?.ElementRef;

            var seen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var verboseList = new System.Collections.Generic.List<string>();
            var baseNames = new System.Collections.Generic.List<string>();

            void CollectParams(Element element, string sourceLabel)
            {
                if (element == null) return;
                foreach (Parameter p in element.Parameters)
                {
                    try
                    {
                        string name = p.Definition?.Name ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(name)) continue;
                        if (seen.Contains(name)) continue;

                        // Determine whether parameter is a shared parameter by trying to cast to ExternalDefinition
                        bool isShared = false;
                        try
                        {
                            var extDef = p.Definition as ExternalDefinition;
                            if (extDef != null) isShared = true;
                        }
                        catch { /* ignore */ }

                        string storage = p.StorageType.ToString();
                        string display = $"{name}  ({sourceLabel}, {storage}{(isShared ? ", Shared" : "")})";

                        seen.Add(name);
                        verboseList.Add(display);
                        baseNames.Add(name);
                    }
                    catch { /* ignore bad param */ }
                }
            }

            // collect type params then instance params (so instance-only show up)
            CollectParams(repType, "Type");
            CollectParams(repInst, "Instance");

            // populate collections: ParameterNames = raw names used by export; ParameterInfo = verbose for tooltips/debug
            var ordered = baseNames.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
            foreach (var n in ordered) ParameterNames.Add(n);

            // verbose (keep order)
            foreach (var v in verboseList.OrderBy(s => s.Split(new[] { "  (" }, StringSplitOptions.None)[0], StringComparer.OrdinalIgnoreCase))
                ParameterInfo.Add(v);

            // geometry preview
            var previewEl = repInst ?? repType;
            if (previewEl != null) AddGeometryPreview(previewEl);
        }

        #endregion

        #region Helpers

        private string GetParameterValue(Parameter param)
        {
            if (param == null) return "";
            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String: return param.AsString() ?? "";
                    case StorageType.Double: return param.AsDouble().ToString(CultureInfo.InvariantCulture);
                    case StorageType.Integer: return param.AsInteger().ToString();
                    case StorageType.ElementId: return param.AsElementId()?.Value.ToString() ?? "";
                    default: return "";
                }
            }
            catch { return ""; }
        }

        private void AddGeometryPreview(Element el)
        {
            GeometryPreview.Clear();

            var lp = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
            if (lp != null && lp.StorageType == StorageType.Double) GeometryPreview.Add($"Length: {UnitUtils.ConvertFromInternalUnits(lp.AsDouble(), UnitTypeId.Meters):F3} m");

            var ap = el.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
            if (ap != null && ap.StorageType == StorageType.Double) GeometryPreview.Add($"Area: {UnitUtils.ConvertFromInternalUnits(ap.AsDouble(), UnitTypeId.SquareMeters):F3} m²");

            var vp = el.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
            if (vp != null && vp.StorageType == StorageType.Double) GeometryPreview.Add($"Volume: {UnitUtils.ConvertFromInternalUnits(vp.AsDouble(), UnitTypeId.CubicMeters):F3} m³");

            var w = el.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
            if (w != null && w.StorageType == StorageType.Double) GeometryPreview.Add($"Width: {UnitUtils.ConvertFromInternalUnits(w.AsDouble(), UnitTypeId.Millimeters):F0} mm");
        }

        private string TryGetElementName(Element el)
        {
            try
            {
                if (!string.IsNullOrEmpty(el.Name)) return el.Name;
                var p = el.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
                if (p != null) return GetParameterValue(p);
            }
            catch { }
            return el.GetType().Name;
        }

        private string TryGetLevelName(Element el)
        {
            try
            {
                if (el.LevelId != null && el.LevelId != ElementId.InvalidElementId)
                {
                    var lvl = doc.GetElement(el.LevelId) as Level;
                    if (lvl != null) return lvl.Name;
                }
                var p = el.get_Parameter(BuiltInParameter.LEVEL_PARAM);
                if (p != null) return GetParameterValue(p);
            }
            catch { }
            return "";
        }

        private string BuildInstanceInfoSummary(Element el)
        {
            try
            {
                var parts = new System.Collections.Generic.List<string>();
                var l = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                var a = el.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                var v = el.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);

                if (l != null && l.StorageType == StorageType.Double) parts.Add($"L={UnitUtils.ConvertFromInternalUnits(l.AsDouble(), UnitTypeId.Meters):F3}m");
                if (a != null && a.StorageType == StorageType.Double) parts.Add($"A={UnitUtils.ConvertFromInternalUnits(a.AsDouble(), UnitTypeId.SquareMeters):F3}m²");
                if (v != null && v.StorageType == StorageType.Double) parts.Add($"V={UnitUtils.ConvertFromInternalUnits(v.AsDouble(), UnitTypeId.CubicMeters):F3}m³");

                if (!parts.Any())
                {
                    var mk = el.LookupParameter("Mark") ?? el.LookupParameter("Type Mark");
                    if (mk != null) parts.Add($"Mark={GetParameterValue(mk)}");
                }

                return string.Join("; ", parts);
            }
            catch { return ""; }
        }

        private string GetRoomName(Element el, Document targetDoc)
        {
            if (el == null) return "";
            try
            {
                // 1. Doors and Windows (Transition logic)
                if (el is FamilyInstance fi)
                {
                    var catId = fi.Category?.Id.Value ?? 0;
                    if (catId == (int)BuiltInCategory.OST_Doors || catId == (int)BuiltInCategory.OST_Windows)
                    {
                        if (fi.FromRoom != null || fi.ToRoom != null)
                        {
                            string from = fi.FromRoom?.Name ?? "External/None";
                            string to = fi.ToRoom?.Name ?? "External/None";
                            return $"{from} to {to}";
                        }
                    }
                    
                    // 2. Family Instances with spatial Room property
                    if (fi.Room != null) return fi.Room.Name;
                }

                // 3. Spatial calculation for elements with physical location (Furniture, Equipment, etc.)
                var phaseId = el.get_Parameter(BuiltInParameter.PHASE_CREATED)?.AsElementId();
                Phase phase = (phaseId != null && phaseId != ElementId.InvalidElementId) ? targetDoc.GetElement(phaseId) as Phase : null;
                if (phase == null) phase = targetDoc.Phases.Cast<Phase>().LastOrDefault(); // Fallback to last phase

                if (phase != null)
                {
                    Location loc = el.Location;
                    if (loc is LocationPoint lp)
                    {
                        var rm = targetDoc.GetRoomAtPoint(lp.Point, phase);
                        if (rm != null) return rm.Name;
                    }
                    else if (loc is LocationCurve lc)
                    {
                        XYZ mid = lc.Curve.Evaluate(0.5, true);
                        var rm = targetDoc.GetRoomAtPoint(mid, phase);
                        if (rm != null) return rm.Name;
                    }
                }

                // 4. Fallback to parameters
                var p = el.get_Parameter(BuiltInParameter.ROOM_NAME) ?? el.get_Parameter(BuiltInParameter.ELEM_ROOM_NAME);
                if (p != null && !string.IsNullOrEmpty(p.AsString())) return p.AsString();
            }
            catch { }
            return "No Room";
        }

        private string EscapeCsv(string s)
        {
            if (s == null) return "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n")) return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        #endregion

        #region Export logic

        public void Export_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedTypes = TypeEntries.Where(t => t.IsSelected).ToList();
                if (!selectedTypes.Any())
                {
                    MessageBox.Show("Please select at least one type to export.", "No types selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                bool itemizeAll = CHK_ItemizeAll?.IsChecked == true;
                bool breakdownByLevel = CHK_BreakdownByLevel?.IsChecked == true;
                bool includeCount = CHK_Count?.IsChecked == true;
                bool includeLength = CHK_Length?.IsChecked == true;
                bool includeArea = CHK_Area?.IsChecked == true;
                bool includeVolume = CHK_Volume?.IsChecked == true;

                int decimals = 2;
                if (CB_Decimals?.SelectedItem is System.Windows.Controls.ComboBoxItem ci)
                {
                    int.TryParse(ci.Content.ToString(), out decimals);
                }

                var saveDlg = new SaveFileDialog
                {
                    Filter = "CSV file (*.csv)|*.csv",
                    FileName = $"RevitExport_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
                    DefaultExt = "csv"
                };
                if (saveDlg.ShowDialog() != true) return;

                string filePath = saveDlg.FileName;
                using (var sw = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    var modeHeader = (TAB_Mode.SelectedItem as System.Windows.Controls.TabItem)?.Header?.ToString() ?? "BOQ";
                    if (modeHeader.StartsWith("BOQ"))
                    {
                        var itemizeTypes = selectedTypes.Where(t => itemizeAll || t.ExportInstances).ToList();
                        var aggregateTypes = selectedTypes.Except(itemizeTypes).ToList();
                        
                        int totalSteps = aggregateTypes.Count + selectedTypes.Where(t => itemizeAll || t.ExportInstances).Sum(t => t.Instances.Count);
                        int currentStep = 0;

                        if (aggregateTypes.Any())
                        {
                            var header = new System.Collections.Generic.List<string> { "Category", "TypeName", "TypeId" };
                            if (breakdownByLevel) header.Add("Level");
                            if (includeCount) header.Add("Quantity");
                            if (includeLength) header.Add("TotalLength_m");
                            if (includeArea) header.Add("TotalArea_m2");
                            if (includeVolume) header.Add("TotalVolume_m3");
                            sw.WriteLine(string.Join(",", header.Select(EscapeCsv)));

                            foreach (var te in aggregateTypes)
                            {
                                currentStep++;
                                UpdateProgress(currentStep, totalSteps, $"Exporting aggregated type {currentStep} of {totalSteps}");

                                // Get elements from all selected categories that match this type
                                var elems = new FilteredElementCollector(doc)
                                    .WhereElementIsNotElementType()
                                    .Where(x => x.GetTypeId() == te.TypeId)
                                    .ToList();

                                if (!elems.Any()) continue;

                                string catName = elems.First().Category?.Name ?? "Unknown";

                                if (breakdownByLevel)
                                {
                                    var groups = elems.GroupBy(el => { var lv = TryGetLevelName(el); return string.IsNullOrEmpty(lv) ? "<Unassigned>" : lv; });
                                    foreach (var g in groups)
                                    {
                                        int qty = g.Count();
                                        double sumLen = 0, sumArea = 0, sumVol = 0;
                                        foreach (var el in g)
                                        {
                                            var lp = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                                            if (lp != null && lp.StorageType == StorageType.Double) sumLen += UnitUtils.ConvertFromInternalUnits(lp.AsDouble(), UnitTypeId.Meters);
                                            var ap = el.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                                            if (ap != null && ap.StorageType == StorageType.Double) sumArea += UnitUtils.ConvertFromInternalUnits(ap.AsDouble(), UnitTypeId.SquareMeters);
                                            var vp = el.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                                            if (vp != null && vp.StorageType == StorageType.Double) sumVol += UnitUtils.ConvertFromInternalUnits(vp.AsDouble(), UnitTypeId.CubicMeters);
                                        }

                                        var row = new System.Collections.Generic.List<string> { catName, te.Name, te.TypeId.Value.ToString(), g.Key };
                                        if (includeCount) row.Add(qty.ToString());
                                        if (includeLength) row.Add(sumLen.ToString($"F{decimals}", CultureInfo.InvariantCulture));
                                        if (includeArea) row.Add(sumArea.ToString($"F{decimals}", CultureInfo.InvariantCulture));
                                        if (includeVolume) row.Add(sumVol.ToString($"F{decimals}", CultureInfo.InvariantCulture));
                                        sw.WriteLine(string.Join(",", row.Select(EscapeCsv)));
                                    }
                                }
                                else
                                {
                                    int qty = elems.Count();
                                    double sumLen = 0, sumArea = 0, sumVol = 0;
                                    foreach (var el in elems)
                                    {
                                        var lp = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                                        if (lp != null && lp.StorageType == StorageType.Double) sumLen += UnitUtils.ConvertFromInternalUnits(lp.AsDouble(), UnitTypeId.Meters);
                                        var ap = el.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                                        if (ap != null && ap.StorageType == StorageType.Double) sumArea += UnitUtils.ConvertFromInternalUnits(ap.AsDouble(), UnitTypeId.SquareMeters);
                                        var vp = el.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                                        if (vp != null && vp.StorageType == StorageType.Double) sumVol += UnitUtils.ConvertFromInternalUnits(vp.AsDouble(), UnitTypeId.CubicMeters);
                                    }

                                    var row = new System.Collections.Generic.List<string> { catName, te.Name, te.TypeId.Value.ToString() };
                                    if (includeCount) row.Add(qty.ToString());
                                    if (includeLength) row.Add(sumLen.ToString($"F{decimals}", CultureInfo.InvariantCulture));
                                    if (includeArea) row.Add(sumArea.ToString($"F{decimals}", CultureInfo.InvariantCulture));
                                    if (includeVolume) row.Add(sumVol.ToString($"F{decimals}", CultureInfo.InvariantCulture));
                                    sw.WriteLine(string.Join(",", row.Select(EscapeCsv)));
                                }
                            }
                        }

                        if (selectedTypes.Any(t => itemizeAll || t.ExportInstances))
                        {
                            sw.WriteLine();
                            var hdr = new System.Collections.Generic.List<string> { "Category", "TypeName", "ElementId", "UniqueId", "Name", "Level", "Room", "Length_m", "Area_m2", "Volume_m3", "InfoSummary" };
                            sw.WriteLine(string.Join(",", hdr.Select(EscapeCsv)));

                            var itemTypes = selectedTypes.Where(t => itemizeAll || t.ExportInstances);
                            foreach (var te in itemTypes)
                            {
                                foreach (var inst in te.Instances)
                                {
                                    currentStep++;
                                    if (currentStep % 10 == 0 || currentStep == totalSteps)
                                        UpdateProgress(currentStep, totalSteps, $"Exporting instance {currentStep} of {totalSteps}");

                                    var el = inst.ElementRef;
                                    double len = 0, area = 0, vol = 0;
                                    var lp = el?.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                                    if (lp != null && lp.StorageType == StorageType.Double) len = UnitUtils.ConvertFromInternalUnits(lp.AsDouble(), UnitTypeId.Meters);
                                    var ap = el?.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                                    if (ap != null && ap.StorageType == StorageType.Double) area = UnitUtils.ConvertFromInternalUnits(ap.AsDouble(), UnitTypeId.SquareMeters);
                                    var vp = el?.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                                    if (vp != null && vp.StorageType == StorageType.Double) vol = UnitUtils.ConvertFromInternalUnits(vp.AsDouble(), UnitTypeId.CubicMeters);

                                    string room = GetRoomName(el, doc);

                                    var row = new System.Collections.Generic.List<string>
                                    {
                                        el?.Category?.Name ?? "Unknown",
                                        te.Name,
                                        inst.Id,
                                        el?.UniqueId ?? "",
                                        inst.Name,
                                        inst.Level,
                                        room,
                                        (len>0?len.ToString($"F{decimals}", CultureInfo.InvariantCulture):""),
                                        (area>0?area.ToString($"F{decimals}", CultureInfo.InvariantCulture):""),
                                        (vol>0?vol.ToString($"F{decimals}", CultureInfo.InvariantCulture):""),
                                        inst.InfoSummary ?? ""
                                    };
                                    sw.WriteLine(string.Join(",", row.Select(EscapeCsv)));
                                }
                            }
                        }
                    }
                    else // QA-QC
                    {
                        var qaTypes = TypeEntries.Where(t => t.IsSelected && t.QAQCEnabled).ToList();
                        if (!qaTypes.Any())
                        {
                            MessageBox.Show("No types selected for QA-QC export. Mark QA-QC for the types you want to export.", "No QA Types", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }

                        var selectedParams = LB_TypeParameters.SelectedItems.Cast<string>().ToList();
                        if (!selectedParams.Any())
                        {
                            MessageBox.Show("Select one or more parameters from the 'Selected Type Parameters' list to include in the QA-QC CSV.", "No parameters", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }

                        int totalSteps = qaTypes.Sum(t => t.Instances.Count);
                        int currentStep = 0;

                        var header = new System.Collections.Generic.List<string> { "Category", "TypeName", "ElementId", "UniqueId", "Name", "Level", "Room" };
                        header.AddRange(selectedParams);
                        sw.WriteLine(string.Join(",", header.Select(EscapeCsv)));

                        foreach (var te in qaTypes)
                        {
                            foreach (var inst in te.Instances)
                            {
                                currentStep++;
                                if (currentStep % 5 == 0 || currentStep == totalSteps)
                                    UpdateProgress(currentStep, totalSteps, $"Exporting QA-QC {currentStep} of {totalSteps}");

                                var el = inst.ElementRef;
                                string room = GetRoomName(el, doc);

                                var row = new System.Collections.Generic.List<string>
                                {
                                    el?.Category?.Name ?? "Unknown",
                                    te.Name,
                                    inst.Id,
                                    el?.UniqueId ?? "",
                                    inst.Name,
                                    inst.Level,
                                    room
                                };

                                foreach (var pName in selectedParams)
                                {
                                    string val = "";
                                    try
                                    {
                                        var p = el.LookupParameter(pName) ?? te.RepresentativeTypeElement?.LookupParameter(pName);
                                        if (p != null) val = GetParameterValue(p);
                                    }
                                    catch { val = ""; }
                                    row.Add(val);
                                }

                                sw.WriteLine(string.Join(",", row.Select(EscapeCsv)));
                            }
                        }
                    }

                    MessageBox.Show($"Export completed:\n{filePath}", "Export OK", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Export failed: " + ex.Message + "\n\nPlease contact support: mohammad.shahnawaz@expocitydubai.ae", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                UpdateProgress(1, 1, "Ready");
            }
        }

        public void Import_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openDlg = new OpenFileDialog
                {
                    Filter = "CSV file (*.csv)|*.csv",
                    Title = "Select CSV file to import"
                };

                if (openDlg.ShowDialog() == true)
                {
                    _eventHandler.FilePath = openDlg.FileName;
                    _eventHandler.IsImport = true;
                    _externalEvent.Raise();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Import trigger failed: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Double-click toggle for parameter list

        private void LB_TypeParameters_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                var list = sender as System.Windows.Controls.ListBox;
                if (list == null) return;

                // Fast path: if an item is selected already, it's the clicked item
                var clicked = list.SelectedItem as string;

                // If clicked is null (double-click on unselected item), hit-test to find the item
                if (clicked == null)
                {
                    var pt = e.GetPosition(list);
                    var element = list.InputHitTest(pt) as System.Windows.DependencyObject;
                    while (element != null && !(element is System.Windows.Controls.ListBoxItem))
                    {
                        element = System.Windows.Media.VisualTreeHelper.GetParent(element);
                    }
                    var lbi = element as System.Windows.Controls.ListBoxItem;
                    if (lbi != null) clicked = lbi.Content as string;
                }

                if (string.IsNullOrEmpty(clicked)) return;

                if (list.SelectedItems.Contains(clicked))
                {
                    list.SelectedItems.Remove(clicked);
                }
                else
                {
                    list.SelectedItems.Add(clicked);
                }

                list.Focus();
                e.Handled = true;
            }
            catch
            {
                try { GeometryPreview.Add("Toggle select failed."); } catch { }
            }
        }

        #endregion

        #region Support mail navigation handler

        private void SupportMail_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
                e.Handled = true;
            }
            catch
            {
                MessageBox.Show("Unable to open email client. Please send an email to mohammad.shahnawaz@expocitydubai.ae", "Contact Support", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        #endregion
    }

    #region Helper classes & converter

    public class InstanceEntry
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Level { get; set; }
        public string InfoSummary { get; set; }
        public Element ElementRef { get; set; }
    }

    public class TypeEntry : INotifyPropertyChanged
    {
        public Element RepresentativeTypeElement { get; set; }
        public ElementId TypeId { get; set; }
        public string Name { get; set; }

        private ObservableCollection<InstanceEntry> _instances = new ObservableCollection<InstanceEntry>();
        public ObservableCollection<InstanceEntry> Instances
        {
            get => _instances;
            set { _instances = value; OnPropertyChanged(nameof(Instances)); OnPropertyChanged(nameof(InstanceCount)); }
        }

        public int InstanceCount => Instances?.Count ?? 0;

        private bool _isSelected = false;
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); }
        }

        private bool _exportInstances = false;
        public bool ExportInstances
        {
            get => _exportInstances;
            set { _exportInstances = value; OnPropertyChanged(nameof(ExportInstances)); }
        }

        private bool _qaqcEnabled = false;
        public bool QAQCEnabled
        {
            get => _qaqcEnabled;
            set { _qaqcEnabled = value; OnPropertyChanged(nameof(QAQCEnabled)); }
        }

        private bool _isVisible = true;
        public bool IsVisible
        {
            get => _isVisible;
            set { _isVisible = value; OnPropertyChanged(nameof(IsVisible)); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string n) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    public static class IconHelper
    {
        [DllImport("user32.dll")]
        static extern int GetWindowLong(IntPtr hwnd, int index);

        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr hwnd, int index, int newStyle);

        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hwnd, IntPtr hwndInsertAfter, int x, int y, int width, int height, uint flags);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);

        const int GWL_EXSTYLE = -20;
        const int WS_EX_DLGMODALFRAME = 0x0001;
        const int SWP_NOSIZE = 0x0001;
        const int SWP_NOMOVE = 0x0002;
        const int SWP_NOZORDER = 0x0004;
        const int SWP_FRAMECHANGED = 0x0020;
        const uint WM_SETICON = 0x0080;

        public static void RemoveIcon(Window window)
        {
            var hwnd = new WindowInteropHelper(window).Handle;
            int extendedStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            SetWindowLong(hwnd, GWL_EXSTYLE, extendedStyle | WS_EX_DLGMODALFRAME);
            SendMessage(hwnd, WM_SETICON, new IntPtr(1), IntPtr.Zero);
            SendMessage(hwnd, WM_SETICON, IntPtr.Zero, IntPtr.Zero);
            SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
        }
    }

    #endregion
}
