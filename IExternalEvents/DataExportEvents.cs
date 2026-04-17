using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CostAnalysis.IExternalEvents
{
    public class DataExportEvents : IExternalEventHandler
    {
        public string FilePath { get; set; }
        public bool IsImport { get; set; }
        public Action<int, int, string> ProgressUpdater { get; set; }

        public void Execute(UIApplication app)
        {
            if (!IsImport || string.IsNullOrEmpty(FilePath) || !File.Exists(FilePath))
                return;

            try
            {
                var doc = app.ActiveUIDocument.Document;
                var lines = File.ReadAllLines(FilePath);
                if (lines.Length < 2) return;

                var headers = ParseCsvLine(lines[0]);
                int uniqueIdIdx = Array.IndexOf(headers, "UniqueId");

                if (uniqueIdIdx == -1)
                {
                    TaskDialog.Show("Import Error", "CSV must contain a 'UniqueId' column.");
                    return;
                }

                int successCount = 0;
                int failCount = 0;
                var report = new System.Text.StringBuilder();
                report.AppendLine("Import Report:");
                report.AppendLine("----------------------------");

                using (Transaction t = new Transaction(doc, "Import Data from CSV"))
                {
                    t.Start();

                    int totalLines = lines.Length - 1;
                    for (int i = 1; i < lines.Length; i++)
                    {
                        if (i % 10 == 0 || i == totalLines)
                            ProgressUpdater?.Invoke(i, totalLines, $"Importing row {i} of {totalLines}");


                        if (string.IsNullOrWhiteSpace(lines[i])) continue;
                        var values = ParseCsvLine(lines[i]);
                        if (values.Length <= uniqueIdIdx) continue;

                        string uid = values[uniqueIdIdx];
                        Element el = doc.GetElement(uid);

                        if (el == null)
                        {
                            failCount++;
                            report.AppendLine($"Error: Element with UniqueId {uid} not found.");
                            continue;
                        }

                        bool updated = false;
                        string elName = el.Name;
                        var elUpdates = new List<string>();

                        for (int j = 0; j < headers.Length; j++)
                        {
                            if (j == uniqueIdIdx) continue;
                            string pName = headers[j];
                            string pValue = values[j];

                            if (pName == "Category" || pName == "TypeName" || pName == "ElementId" || pName == "Name" || pName == "Level")
                                continue;

                            Parameter p = el.LookupParameter(pName);
                            if (p != null && !p.IsReadOnly)
                            {
                                try
                                {
                                    string oldValue = GetParamValueString(p);
                                    if (oldValue == pValue) continue;

                                    if (p.StorageType == StorageType.String)
                                        p.Set(pValue);
                                    else if (p.StorageType == StorageType.Double)
                                        p.Set(double.Parse(pValue));
                                    else if (p.StorageType == StorageType.Integer)
                                        p.Set(int.Parse(pValue));

                                    elUpdates.Add($"{pName}: {oldValue} -> {pValue}");
                                    updated = true;
                                }
                                catch { /* skip */ }
                            }
                        }

                        if (updated)
                        {
                            successCount++;
                            report.AppendLine($"Updated Element [{el.Id}]: {elName}");
                            foreach (var u in elUpdates) report.AppendLine($"  - {u}");
                        }
                    }

                    t.Commit();
                }

                report.AppendLine("----------------------------");
                report.AppendLine($"Summary: {successCount} elements updated, {failCount} failed.");

                ProgressUpdater?.Invoke(1, 1, "Ready");

                TaskDialog mainDialog = new TaskDialog("Import Complete")
                {
                    MainInstruction = $"Updated {successCount} elements.",
                    MainContent = report.Length > 2000 ? report.ToString().Substring(0, 1997) + "..." : report.ToString(),
                    CommonButtons = TaskDialogCommonButtons.Ok
                };
                mainDialog.Show();
            }
            catch (Exception ex)
            {
                ProgressUpdater?.Invoke(1, 1, "Ready");
                TaskDialog.Show("Import Exception", ex.ToString());
            }
        }

        private string GetParamValueString(Parameter p)
        {
            if (p == null) return "";
            switch (p.StorageType)
            {
                case StorageType.String: return p.AsString() ?? "";
                case StorageType.Double: return p.AsDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
                case StorageType.Integer: return p.AsInteger().ToString();
                case StorageType.ElementId: return p.AsElementId()?.Value.ToString() ?? "";
                default: return "";
            }
        }

        private string[] ParseCsvLine(string line)
        {
            var result = new List<string>();
            bool inQuotes = false;
            var current = new System.Text.StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '\"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(current.ToString().Trim('\"'));
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }
            result.Add(current.ToString().Trim('\"'));
            return result.ToArray();
        }

        public string GetName() => "Data Export Event Handler";
    }
}
