using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using OfficeOpenXml;
using Autodesk.Revit.DB;
using ScheduleLink.Helpers;
using ScheduleLink.Models;

namespace ScheduleLink.Services
{
    public static class ExcelImportService
    {
        public static ScheduleExportData ReadExcelFile(string filePath)
        {
            var data = new ScheduleExportData();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Find data sheet (skip Instructions)
                ExcelWorksheet ws = null;
                foreach (var sheet in package.Workbook.Worksheets)
                {
                    if (!sheet.Name.Equals("Instructions", StringComparison.OrdinalIgnoreCase))
                    { ws = sheet; break; }
                }

                if (ws == null || ws.Dimension == null) return null;

                data.ScheduleName = ws.Name;
                int rowCount = ws.Dimension.End.Row;
                int colCount = ws.Dimension.End.Column;
                if (rowCount < 2 || colCount < 2) return null;

                // Find Element ID column
                int idCol = FindIdColumn(ws, colCount);

                // Read columns
                var paramCols = new List<int>();
                for (int c = 1; c <= colCount; c++)
                {
                    if (c == idCol) continue;
                    string header = ws.Cells[1, c].Text ?? "";
                    header = header.Trim();
                    if (string.IsNullOrEmpty(header)) continue;

                    bool isReadOnly = DetectReadOnly(ws, c, rowCount);

                    data.Columns.Add(new ScheduleColumnInfo
                    {
                        Name = header,
                        HeaderText = header,
                        IsReadOnly = isReadOnly,
                        FieldIndex = c
                    });
                    paramCols.Add(c);
                }

                // Read data rows
                for (int row = 2; row <= rowCount; row++)
                {
                    string idText = ws.Cells[row, idCol].Text ?? "";
                    idText = idText.Trim();
                    if (string.IsNullOrEmpty(idText) || !long.TryParse(idText, out long elemId))
                        continue;

                    var rowData = new ScheduleRowData { ElementId = elemId };
                    foreach (int c in paramCols)
                    {
                        string val = ws.Cells[row, c].Text ?? "";
                        rowData.Values.Add(val.Trim());
                    }
                    data.Rows.Add(rowData);
                }

                ReadScheduleIdFromInstructions(package, data);
            }

            return data;
        }

        public static ImportResult ImportToRevit(Document doc, ScheduleExportData excelData)
        {
            var result = new ImportResult { TotalRows = excelData.Rows.Count };

            // Create progress window
            var progressWin = new System.Windows.Window
            {
                Title = "ScheduleLink - Importing...",
                Width = 420,
                Height = 120,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                ResizeMode = System.Windows.ResizeMode.NoResize,
                WindowStyle = System.Windows.WindowStyle.ToolWindow,
                Topmost = true
            };

            var stack = new System.Windows.Controls.StackPanel { Margin = new System.Windows.Thickness(16) };
            var statusText = new System.Windows.Controls.TextBlock
            {
                Text = "Preparing import...",
                FontSize = 13,
                Margin = new System.Windows.Thickness(0, 0, 0, 8)
            };
            var progressBar = new System.Windows.Controls.ProgressBar
            {
                Height = 22,
                Minimum = 0,
                Maximum = excelData.Rows.Count
            };
            stack.Children.Add(statusText);
            stack.Children.Add(progressBar);
            progressWin.Content = stack;
            progressWin.Show();

            using (var tx = new Transaction(doc, "ScheduleLink Import"))
            {
                tx.Start();
                try
                {
                    for (int rowIdx = 0; rowIdx < excelData.Rows.Count; rowIdx++)
                    {
                        var row = excelData.Rows[rowIdx];

                        // Update progress every 10 rows
                        if (rowIdx % 10 == 0)
                        {
                            progressBar.Value = rowIdx;
                            statusText.Text = "Importing row " + (rowIdx + 1) + " of " + excelData.Rows.Count + "...";
                            // Force UI update
                            System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(
                                System.Windows.Threading.DispatcherPriority.Background,
                                new System.Action(delegate { }));
                        }

                        ElementId eid = row.ElementId.ToElementId();
                        Element elem = doc.GetElement(eid);
                        if (elem == null)
                        {
                            result.ElementsNotFound++;
                            if (result.Errors.Count < 10)
                                result.Errors.Add("Element " + row.ElementId + " not found");
                            continue;
                        }

                        for (int c = 0; c < excelData.Columns.Count && c < row.Values.Count; c++)
                        {
                            var colInfo = excelData.Columns[c];
                            if (colInfo.IsReadOnly || colInfo.IsCalculated)
                            {
                                result.SkippedReadOnly++;
                                continue;
                            }

                            string newValue = row.Values[c];

                            Parameter param = ScheduleReaderService.FindParameter(elem, colInfo.HeaderText);
                            if (param == null && colInfo.Name != colInfo.HeaderText)
                                param = ScheduleReaderService.FindParameter(elem, colInfo.Name);

                            if (param == null)
                            {
                                result.ParamNotFound++;
                                continue;
                            }

                            if (param.IsReadOnly)
                            {
                                result.SkippedReadOnly++;
                                continue;
                            }

                            string currentValue = GetParameterDisplayValue(param);
                            if (currentValue == newValue)
                            {
                                result.SkippedUnchanged++;
                                continue;
                            }

                            if (SetParameterValue(param, newValue))
                                result.UpdatedParams++;
                            else
                                result.FailedParams++;
                        }
                    }

                    tx.Commit();
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    result.Errors.Add("Transaction failed: " + ex.Message);
                }
            }

            progressWin.Close();
            return result;
        }

        #region Private Helpers

        private static int FindIdColumn(ExcelWorksheet ws, int colCount)
        {
            for (int c = 1; c <= colCount; c++)
            {
                string h = ws.Cells[1, c].Text ?? "";
                if (h.Trim().Equals("Element ID", StringComparison.OrdinalIgnoreCase))
                    return c;
            }
            return 1;
        }

        private static bool DetectReadOnly(ExcelWorksheet ws, int col, int rowCount)
        {
            if (rowCount < 2) return false;
            var bg = ws.Cells[2, col].Style.Fill.BackgroundColor;
            if (bg == null || string.IsNullOrEmpty(bg.Rgb)) return false;

            string rgb = bg.Rgb.ToUpper();
            return rgb.Contains("FFC7CE") || rgb.Contains("FF0000") ||
                   rgb.Contains("D9D9D9") || rgb.Contains("C0C0C0");
        }

        private static void ReadScheduleIdFromInstructions(ExcelPackage package, ScheduleExportData data)
        {
            try
            {
                var instrWs = package.Workbook.Worksheets["Instructions"];
                if (instrWs == null || instrWs.Dimension == null) return;

                for (int r = 1; r <= Math.Min(instrWs.Dimension.End.Row, 20); r++)
                {
                    string label = instrWs.Cells[r, 1].Text ?? "";
                    if (label.Trim().StartsWith("Schedule ID"))
                    {
                        string val = instrWs.Cells[r, 2].Text ?? "";
                        if (long.TryParse(val.Trim(), out long id))
                            data.ScheduleId = id;
                        break;
                    }
                }
            }
            catch { }
        }

        private static string GetParameterDisplayValue(Parameter param)
        {
            if (param == null || !param.HasValue)
                return string.Empty;

            switch (param.StorageType)
            {
                case StorageType.String:
                    return param.AsString() ?? string.Empty;

                case StorageType.Integer:
                    return param.AsInteger().ToString();

                case StorageType.Double:
                    return param.AsValueString() ?? param.AsDouble().ToString(CultureInfo.InvariantCulture);

                case StorageType.ElementId:
                    return param.AsElementId().GetIdValueLong().ToString();

                default:
                    return string.Empty;
            }
        }

        private static bool SetParameterValue(Parameter param, string value)
        {
            if (param == null || param.IsReadOnly) return false;

            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        param.Set(value ?? "");
                        return true;

                    case StorageType.Integer:
                        if (int.TryParse(value, out int intVal))
                        { param.Set(intVal); return true; }

                        string lower = (value ?? "").Trim().ToLower();
                        if (lower == "yes" || lower == "כן" || lower == "true" || lower == "1")
                        { param.Set(1); return true; }
                        if (lower == "no" || lower == "לא" || lower == "false" || lower == "0")
                        { param.Set(0); return true; }
                        return false;

                    case StorageType.Double:
                        if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double d))
                        { param.Set(ConvertToInternal(param, d)); return true; }
                        if (double.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out double d2))
                        { param.Set(ConvertToInternal(param, d2)); return true; }
                        return false;

                    case StorageType.ElementId:
                        if (long.TryParse(value, out long eidVal))
                        { param.Set(eidVal.ToElementId()); return true; }
                        return false;

                    default:
                        return false;
                }
            }
            catch { return false; }
        }

        private static double ConvertToInternal(Parameter param, double displayValue)
        {
            try
            {
                // Use reflection to handle both old and new unit APIs
                var getUnitTypeId = typeof(Parameter).GetMethod("GetUnitTypeId");
                if (getUnitTypeId != null)
                {
                    // Revit 2021+ (ForgeTypeId)
                    var unitTypeId = getUnitTypeId.Invoke(param, null);
                    if (unitTypeId != null)
                    {
                        var convertMethod = typeof(UnitUtils).GetMethod("ConvertToInternalUnits",
                            new[] { typeof(double), unitTypeId.GetType() });
                        if (convertMethod != null)
                        {
                            var result = convertMethod.Invoke(null, new object[] { displayValue, unitTypeId });
                            return (double)result;
                        }
                    }
                }
            }
            catch { }

            return displayValue;
        }

        #endregion
    }
}
