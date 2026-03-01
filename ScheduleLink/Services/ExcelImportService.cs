// ExcelImportService.cs

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

        // ============================================================
        // REPLACE the entire ImportToRevit method in ExcelImportService.cs
        // AND add the two helper classes at the end of the file
        // (before the last closing } of the namespace)
        // ============================================================

        // --- METHOD: Replace ImportToRevit ---

        public static ImportResult ImportToRevit(Document doc, ScheduleExportData excelData)
        {
            var result = new ImportResult { TotalRows = excelData.Rows.Count };

            Logger.Info(Logger.LogCategory.Import, "ImportToRevit: Starting - Rows=" + excelData.Rows.Count + " Cols=" + excelData.Columns.Count);

            // --- Validation: check for invalid Element IDs ---
            int invalidIds = 0;
            foreach (var row in excelData.Rows)
            {
                if (row.ElementId <= 0)
                    invalidIds++;
            }
            if (invalidIds > 0)
            {
                Logger.Info(Logger.LogCategory.Import, "ImportToRevit: WARNING - " + invalidIds + " rows with invalid Element IDs");
                result.Errors.Add(invalidIds + " rows have invalid Element IDs");
            }

            // --- Validation: warn if read-only columns have changed values ---
            foreach (var col in excelData.Columns)
            {
                if (col.IsReadOnly)
                    Logger.Info(Logger.LogCategory.Import, "  Column '" + col.HeaderText + "' = Read-only (will skip)");
            }

            using (var tx = new Transaction(doc, "ScheduleLink Import"))
            {
                tx.Start();

                // Set failure handler to continue on errors
                var failureHandler = new IgnoreWarningsHandler();
                failureHandler.Doc = doc;
                var failOpts = tx.GetFailureHandlingOptions();
                failOpts.SetFailuresPreprocessor(failureHandler);
                tx.SetFailureHandlingOptions(failOpts);

                try
                {
                    // === PASS 1: Identify unique parameters that need two-pass handling ===
                    var uniqueParamChanges = new List<UniqueParamChange>();

                    for (int rowIdx = 0; rowIdx < excelData.Rows.Count; rowIdx++)
                    {
                        var row = excelData.Rows[rowIdx];
                        ElementId eid = row.ElementId.ToElementId();
                        Element elem = doc.GetElement(eid);
                        if (elem == null) continue;

                        for (int c = 0; c < excelData.Columns.Count && c < row.Values.Count; c++)
                        {
                            var colInfo = excelData.Columns[c];
                            if (colInfo.IsReadOnly || colInfo.IsCalculated) continue;

                            string newValue = row.Values[c];
                            Parameter param = ScheduleReaderService.FindParameter(elem, colInfo.HeaderText);
                            if (param == null && colInfo.Name != colInfo.HeaderText)
                                param = ScheduleReaderService.FindParameter(elem, colInfo.Name);
                            if (param == null || param.IsReadOnly) continue;

                            string currentValue = GetParameterDisplayValue(param);
                            if (currentValue == newValue) continue;

                            // Check if this is a unique parameter (only Sheet Number is blocked by Revit)
                            string paramName = param.Definition.Name;
                            bool isUnique = paramName == "Sheet Number";

                            if (isUnique)
                            {
                                uniqueParamChanges.Add(new UniqueParamChange
                                {
                                    RowIndex = rowIdx,
                                    Element = elem,
                                    Param = param,
                                    ColInfo = colInfo,
                                    NewValue = newValue,
                                    CurrentValue = currentValue
                                });
                            }
                        }
                    }

                    // Set unique params to temp values first to avoid conflicts
                    if (uniqueParamChanges.Count > 0)
                    {
                        Logger.Info(Logger.LogCategory.Import, "Pass 1: Setting " + uniqueParamChanges.Count + " unique params to temp values");
                        for (int i = 0; i < uniqueParamChanges.Count; i++)
                        {
                            var change = uniqueParamChanges[i];
                            string tempValue = "TEMP_SL_" + i + "_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                            try
                            {
                                SetParameterValue(change.Param, tempValue);
                            }
                            catch { }
                        }
                    }

                    // === PASS 2: Set unique params to final values ===
                    if (uniqueParamChanges.Count > 0)
                    {
                        Logger.Info(Logger.LogCategory.Import, "Pass 2: Setting " + uniqueParamChanges.Count + " unique params to final values");
                        foreach (var change in uniqueParamChanges)
                        {
                            try
                            {
                                if (SetParameterValue(change.Param, change.NewValue))
                                    result.UpdatedParams++;
                                else
                                {
                                    result.FailedParams++;
                                    if (result.Errors.Count < 50)
                                        result.Errors.Add("Row " + (change.RowIndex + 2) + ": Failed to set '" + change.ColInfo.HeaderText + "' = '" + change.NewValue + "'");
                                }
                            }
                            catch (Exception ex)
                            {
                                result.FailedParams++;
                                if (result.Errors.Count < 50)
                                    result.Errors.Add("Row " + (change.RowIndex + 2) + ": " + change.ColInfo.HeaderText + " - " + ex.Message);
                            }
                        }
                    }

                    // === PASS 3: Set all other (non-unique) parameters ===
                    var handledKeys = new HashSet<string>();
                    foreach (var change in uniqueParamChanges)
                        handledKeys.Add(change.RowIndex + "_" + change.ColInfo.HeaderText);

                    for (int rowIdx = 0; rowIdx < excelData.Rows.Count; rowIdx++)
                    {
                        var row = excelData.Rows[rowIdx];

                        try
                        {
                            ElementId eid = row.ElementId.ToElementId();
                            Element elem = doc.GetElement(eid);
                            if (elem == null)
                            {
                                result.ElementsNotFound++;
                                if (result.Errors.Count < 50)
                                    result.Errors.Add("Row " + (rowIdx + 2) + ": Element " + row.ElementId + " not found");
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

                                // Skip if already handled as unique param
                                string key = rowIdx + "_" + colInfo.HeaderText;
                                if (handledKeys.Contains(key))
                                    continue;

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

                                try
                                {
                                    if (SetParameterValue(param, newValue))
                                        result.UpdatedParams++;
                                    else
                                    {
                                        result.FailedParams++;
                                        if (result.Errors.Count < 50)
                                            result.Errors.Add("Row " + (rowIdx + 2) + ": Failed to set '" + colInfo.HeaderText + "' = '" + newValue + "'");
                                    }
                                }
                                catch (Exception setEx)
                                {
                                    result.FailedParams++;
                                    if (result.Errors.Count < 50)
                                        result.Errors.Add("Row " + (rowIdx + 2) + ": " + colInfo.HeaderText + " - " + setEx.Message);
                                }
                            }
                        }
                        catch (Exception rowEx)
                        {
                            result.FailedParams++;
                            if (result.Errors.Count < 50)
                                result.Errors.Add("Row " + (rowIdx + 2) + ": " + rowEx.Message);
                        }
                    }

                    // Commit
                    try
                    {
                        tx.Commit();

                        if (failureHandler.FailureMessages.Count > 0)
                        {
                            result.FailedParams += failureHandler.FailureMessages.Count;
                            foreach (var msg in failureHandler.FailureMessages)
                            {
                                if (result.Errors.Count < 50)
                                    result.Errors.Add(msg);
                            }
                            Logger.Info(Logger.LogCategory.Import, "Revit failures: " + failureHandler.FailureMessages.Count);
                        }

                        Logger.Info(Logger.LogCategory.Import, "ImportToRevit: Committed - Updated=" + result.UpdatedParams + " Failed=" + result.FailedParams);
                    }
                    catch (Exception commitEx)
                    {
                        Logger.Error(Logger.LogCategory.Import, "ImportToRevit: Commit failed", commitEx);
                        result.Errors.Add("Commit error: " + commitEx.Message);
                        result.FailedParams++;
                    }
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    Logger.Error(Logger.LogCategory.Import, "ImportToRevit: Transaction failed", ex);
                    result.Errors.Add("Transaction failed: " + ex.Message);
                }
            }

            return result;
        }


        // ============================================================
        // ADD these two classes at the END of ExcelImportService.cs
        // (before the last closing } of the namespace)
        // ============================================================

        /// <summary>
        /// Tracks a unique parameter change for two-pass processing.
        /// </summary>
        internal class UniqueParamChange
        {
            public int RowIndex { get; set; }
            public Element Element { get; set; }
            public Parameter Param { get; set; }
            public ScheduleColumnInfo ColInfo { get; set; }
            public string NewValue { get; set; }
            public string CurrentValue { get; set; }
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

    internal class IgnoreWarningsHandler : IFailuresPreprocessor
    {
        public List<string> FailureMessages { get; } = new List<string>();
        public Document Doc { get; set; }

        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            var failures = failuresAccessor.GetFailureMessages();
            foreach (var failure in failures)
            {
                var severity = failure.GetSeverity();
                string desc = failure.GetDescriptionText();

                var elementIds = failure.GetAdditionalElementIds();
                string elemInfo = "";
                if (elementIds != null && elementIds.Count > 0 && Doc != null)
                {
                    var idStrings = new List<string>();
                    foreach (var eid in elementIds)
                    {
                        Element elem = Doc.GetElement(eid);
                        string name = elem != null
                            ? elem.Name + " (id " + eid.ToString() + ")"
                            : eid.ToString();
                        idStrings.Add(name);
                    }
                    elemInfo = " | " + string.Join(", ", idStrings);
                }

                if (severity == FailureSeverity.Warning)
                {
                    FailureMessages.Add("[Warning] " + desc + elemInfo);
                    failuresAccessor.DeleteWarning(failure);
                }
                else if (severity == FailureSeverity.Error)
                {
                    FailureMessages.Add("[Error] " + desc + elemInfo);
                    failuresAccessor.ResolveFailure(failure);
                }
            }
            return FailureProcessingResult.Continue;
        }
    }
}
