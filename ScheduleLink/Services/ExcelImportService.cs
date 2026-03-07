// ExcelImportService.cs

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
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
                ReadColumnMapping(package, data);
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
                    // === PASS 1: Collect ALL Sheet Number params (changed AND unchanged) ===
                    var allUniqueParams = new List<UniqueParamChange>();

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

                            Parameter param = ScheduleReaderService.FindParameter(elem, colInfo.HeaderText);
                            if (param == null && colInfo.Name != colInfo.HeaderText)
                                param = ScheduleReaderService.FindParameter(elem, colInfo.Name);
                            if (param == null || param.IsReadOnly) continue;

                            string paramName = param.Definition.Name;
                            if (paramName == "Sheet Number")
                            {
                                string currentValue = GetParameterDisplayValue(param);
                                string newValue = row.Values[c];

                                allUniqueParams.Add(new UniqueParamChange
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
                    if (allUniqueParams.Count > 0)
                    {
                        Logger.Info(Logger.LogCategory.Import, "Pass 1: Setting " + allUniqueParams.Count + " unique params to temp values");
                        for (int i = 0; i < allUniqueParams.Count; i++)
                        {
                            var change = allUniqueParams[i];
                            string tempValue = "TEMP_SL_" + i + "_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                            try
                            {
                                SetParameterValue(change.Param, tempValue);
                            }
                            catch { }
                        }
                    }

                    // === PASS 2: Set unique params to final values ===
                    if (allUniqueParams.Count > 0)
                    {
                        // Detect duplicates in final values BEFORE applying
                        var valueCounts = new Dictionary<string, int>();
                        foreach (var change in allUniqueParams)
                        {
                            if (!valueCounts.ContainsKey(change.NewValue))
                                valueCounts[change.NewValue] = 0;
                            valueCounts[change.NewValue]++;
                        }

                        Logger.Info(Logger.LogCategory.Import, "Pass 2: Setting " + allUniqueParams.Count + " Sheet Numbers to final values");
                        foreach (var change in allUniqueParams)
                        {
                            try
                            {
                                // If duplicate detected, restore original value
                                if (valueCounts[change.NewValue] > 1)
                                {
                                    SetParameterValue(change.Param, change.CurrentValue);
                                    result.FailedParams++;
                                    if (result.Errors.Count < 200)
                                        result.Errors.Add("Row " + (change.RowIndex + 2) + ": Excel value '" + change.NewValue + "' is duplicate → kept Revit value '" + change.CurrentValue + "'");
                                }
                                else
                                {
                                    if (SetParameterValue(change.Param, change.NewValue))
                                    {
                                        if (change.CurrentValue != change.NewValue)
                                            result.UpdatedParams++;
                                    }
                                    else
                                    {
                                        SetParameterValue(change.Param, change.CurrentValue);
                                        result.FailedParams++;
                                        if (result.Errors.Count < 200)
                                            result.Errors.Add("Row " + (change.RowIndex + 2) + ": Failed to set '" + change.ColInfo.HeaderText + "' = '" + change.NewValue + "'");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                try { SetParameterValue(change.Param, change.CurrentValue); } catch { }
                                result.FailedParams++;
                                if (result.Errors.Count < 200)
                                    result.Errors.Add("Row " + (change.RowIndex + 2) + ": " + change.ColInfo.HeaderText + " - " + ex.Message);
                            }
                        }
                    }

                    // === PASS 3: Set all other (non-unique) parameters ===
                    var handledKeys = new HashSet<string>();
                    foreach (var change in allUniqueParams)
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
                                if (result.Errors.Count < 200)
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
                                        if (result.Errors.Count < 200)
                                            result.Errors.Add("Row " + (rowIdx + 2) + ": Failed to set '" + colInfo.HeaderText + "' = '" + newValue + "'");
                                    }
                                }
                                catch (Exception setEx)
                                {
                                    result.FailedParams++;
                                    if (result.Errors.Count < 200)
                                        result.Errors.Add("Row " + (rowIdx + 2) + ": " + colInfo.HeaderText + " - " + setEx.Message);
                                }
                            }
                        }
                        catch (Exception rowEx)
                        {
                            result.FailedParams++;
                            if (result.Errors.Count < 200)
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
                                if (result.Errors.Count < 200)
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

        private static void ReadColumnMapping(ExcelPackage package, ScheduleExportData data)
        {
            try
            {
                var instrWs = package.Workbook.Worksheets["Instructions"];
                if (instrWs == null || instrWs.Dimension == null) return;

                // Find "Column Mapping:" row
                int startRow = -1;
                for (int r = 1; r <= Math.Min(instrWs.Dimension.End.Row, 40); r++)
                {
                    string label = instrWs.Cells[r, 1].Text ?? "";
                    if (label.Trim().StartsWith("Column Mapping"))
                    { startRow = r + 2; break; }  // skip header row
                }
                if (startRow < 0) return;

                // Build header->name mapping
                var mapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int r = startRow; r <= instrWs.Dimension.End.Row; r++)
                {
                    string header = (instrWs.Cells[r, 1].Text ?? "").Trim();
                    string paramName = (instrWs.Cells[r, 2].Text ?? "").Trim();
                    if (string.IsNullOrEmpty(header)) break;
                    if (!string.IsNullOrEmpty(paramName))
                        mapping[header] = paramName;
                }

                // Apply mapping to columns
                foreach (var col in data.Columns)
                {
                    if (mapping.TryGetValue(col.HeaderText, out string name))
                        col.Name = name;
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
                    // Use AsValueString first - this returns display text like "Fine", "Architectural"
                    // Falls back to raw integer if no display string available
                    string intDisplay = param.AsValueString();
                    if (!string.IsNullOrEmpty(intDisplay))
                        return intDisplay;
                    return param.AsInteger().ToString();

                case StorageType.Double:
                    return param.AsValueString() ?? param.AsDouble().ToString(CultureInfo.InvariantCulture);

                case StorageType.ElementId:
                    // Try to get element name for display (e.g., View Template name, Scope Box name)
                    ElementId eid = param.AsElementId();
                    if (eid != null && eid != ElementId.InvalidElementId)
                    {
                        Element refElem = param.Element?.Document?.GetElement(eid);
                        if (refElem != null && !string.IsNullOrEmpty(refElem.Name))
                            return refElem.Name;
                    }
                    string eidDisplay = param.AsValueString();
                    if (!string.IsNullOrEmpty(eidDisplay))
                        return eidDisplay;
                    return eid.GetIdValueLong().ToString();

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
                        // Try direct integer parse first
                        if (int.TryParse(value, out int intVal))
                        { param.Set(intVal); return true; }

                        // Try Yes/No/boolean variants
                        string lower = (value ?? "").Trim().ToLower();
                        if (lower == "yes" || lower == "\u05DB\u05DF" || lower == "true" || lower == "1")
                        { param.Set(1); return true; }
                        if (lower == "no" || lower == "\u05DC\u05D0" || lower == "false" || lower == "0")
                        { param.Set(0); return true; }

                        // Try known Revit enum mappings (display name -> integer value)
                        int? knownEnum = ResolveKnownEnumValue(param, value);
                        if (knownEnum.HasValue)
                        { param.Set(knownEnum.Value); return true; }

                        // Try SetValueString - handles enum display names generically
                        if (TrySetValueString(param, value))
                            return true;

                        Logger.Info(Logger.LogCategory.Import,
                            $"  SetParameterValue FAILED: param='{param.Definition.Name}' storage=Integer value='{value}' builtIn={GetBuiltInParamName(param)}");
                        return false;

                    case StorageType.Double:
                        // Try SetValueString first - handles unit-formatted strings like 14'-0"
                        if (TrySetValueString(param, value))
                            return true;

                        if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double d))
                        { param.Set(ConvertToInternal(param, d)); return true; }
                        if (double.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out double d2))
                        { param.Set(ConvertToInternal(param, d2)); return true; }
                        return false;

                    case StorageType.ElementId:
                        // Try direct ID parse first
                        if (long.TryParse(value, out long eidVal))
                        { param.Set(eidVal.ToElementId()); return true; }

                        // Try SetValueString
                        if (TrySetValueString(param, value))
                            return true;

                        // Try to find element by name (for View Template, Scope Box, etc.)
                        if (!string.IsNullOrEmpty(value) && value != "<None>" && value != "None" && value != "---")
                        {
                            Document doc = param.Element?.Document;
                            if (doc != null)
                            {
                                Element found = FindElementByName(doc, value);
                                if (found != null)
                                {
                                    param.Set(found.Id);
                                    return true;
                                }
                            }
                        }

                        // Handle "<None>" / "None" - set to InvalidElementId
                        if (string.IsNullOrEmpty(value) || value == "<None>" || value == "None" || value == "---")
                        {
                            param.Set(ElementId.InvalidElementId);
                            return true;
                        }

                        return false;

                    default:
                        return false;
                }
            }
            catch { return false; }
        }

        /// <summary>
        /// Maps known Revit enum display names to their integer values.
        /// This handles BuiltIn parameters like Detail Level, Discipline, etc.
        /// </summary>
        private static int? ResolveKnownEnumValue(Parameter param, string displayValue)
        {
            if (param == null || string.IsNullOrEmpty(displayValue)) return null;

            string trimmed = displayValue.Trim();

            // Get the BuiltIn parameter ID to determine the enum type
            string builtInName = GetBuiltInParamName(param);

            // ViewDetailLevel: Coarse=1, Medium=2, Fine=3
            if (builtInName == "VIEW_DETAIL_LEVEL" ||
                param.Definition.Name == "Detail Level")
            {
                switch (trimmed.ToLower())
                {
                    case "coarse": return 1;
                    case "medium": return 2;
                    case "fine": return 3;
                }
            }

            // ViewDiscipline
            if (builtInName == "VIEW_DISCIPLINE" ||
                param.Definition.Name == "Discipline")
            {
                switch (trimmed.ToLower())
                {
                    case "architectural": return 1;
                    case "structural": return 2;
                    case "mechanical": return 4;
                    case "electrical": return 8;
                    case "plumbing": return 16;
                    case "coordination": return 4095;
                }
            }

            // Generic fallback: try setting values 0-20 and check AsValueString match
            try
            {
                int currentVal = param.AsInteger();
                for (int testVal = 0; testVal <= 20; testVal++)
                {
                    if (testVal == currentVal) continue;
                    param.Set(testVal);
                    string testDisplay = param.AsValueString();
                    if (string.Equals(testDisplay, trimmed, StringComparison.OrdinalIgnoreCase))
                        return testVal;  // already set to correct value
                }
                // Restore original value if no match found
                param.Set(currentVal);
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Gets the BuiltInParameter name for logging and identification
        /// </summary>
        private static string GetBuiltInParamName(Parameter param)
        {
            try
            {
                if (param?.Definition is InternalDefinition intDef)
                    return intDef.BuiltInParameter.ToString();
            }
            catch { }
            return "UNKNOWN";
        }

        /// <summary>
        /// Safely tries param.SetValueString(). Works across all Revit versions.
        /// Verifies the value actually changed after setting.
        /// </summary>
        private static bool TrySetValueString(Parameter param, string value)
        {
            if (param == null || string.IsNullOrEmpty(value)) return false;

            try
            {
                // Remember current value to verify change
                string before = param.AsValueString() ?? "";

                // SetValueString is void in some Revit versions, bool in others
                // Using reflection to handle both cases
                var method = typeof(Parameter).GetMethod("SetValueString", new[] { typeof(string) });
                if (method == null) return false;

                var returnType = method.ReturnType;
                var result = method.Invoke(param, new object[] { value });

                // If method returns bool, use it
                if (returnType == typeof(bool) && result is bool boolResult)
                    return boolResult;

                // If method is void, check if value actually changed
                string after = param.AsValueString() ?? "";
                if (after != before)
                    return true;

                // Check if the new value matches what we tried to set
                // (handles case where before and after are same because value was already correct)
                if (string.Equals(after, value, StringComparison.OrdinalIgnoreCase))
                    return true;

                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Tries to find an element by name in the document.
        /// Searches views (for View Template), scope boxes, and other named elements.
        /// </summary>
        private static Element FindElementByName(Document doc, string name)
        {
            if (doc == null || string.IsNullOrEmpty(name)) return null;

            try
            {
                // Search all elements with a name
                var collector = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType();

                foreach (Element elem in collector)
                {
                    try
                    {
                        if (elem.Name == name)
                            return elem;
                    }
                    catch { }
                }
            }
            catch { }

            return null;
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
            ScheduleLink.Helpers.Logger.Info(
                ScheduleLink.Helpers.Logger.LogCategory.Import,
                "FailureHandler: " + failures.Count + " failure messages received");

            // PASS 1: Collect ALL messages first before any resolve/delete
            var warningFailures = new List<FailureMessageAccessor>();
            var errorFailures = new List<FailureMessageAccessor>();

            foreach (var failure in failures)
            {
                var severity = failure.GetSeverity();
                string desc = failure.GetDescriptionText();
                string severityTag = severity == FailureSeverity.Warning ? "[Warning]" : "[Error]";

                // Collect ALL related elements (failing + additional)
                var allElementIds = new List<ElementId>();
                int failingCount = 0, additionalCount = 0;

                try
                {
                    var failingIds = failure.GetFailingElementIds();
                    if (failingIds != null)
                    {
                        failingCount = failingIds.Count;
                        allElementIds.AddRange(failingIds);
                    }
                }
                catch { }

                try
                {
                    var additionalIds = failure.GetAdditionalElementIds();
                    if (additionalIds != null)
                    {
                        additionalCount = additionalIds.Count;
                        allElementIds.AddRange(additionalIds);
                    }
                }
                catch { }

                ScheduleLink.Helpers.Logger.Info(
                    ScheduleLink.Helpers.Logger.LogCategory.Import,
                    $"  Failure: '{desc}' | Failing={failingCount} Additional={additionalCount} Total={allElementIds.Count}");

                // Create ONE message PER element for clear reporting
                if (allElementIds.Count > 0 && Doc != null)
                {
                    foreach (var eid in allElementIds)
                    {
                        try
                        {
                            Element elem = Doc.GetElement(eid);
                            if (elem != null)
                            {
                                string category = elem.Category?.Name ?? "";
                                string name = elem.Name ?? "";
                                string viewType = "";

                                if (elem is Autodesk.Revit.DB.View view)
                                {
                                    try
                                    {
                                        viewType = view.ViewType.ToString().Replace("FloorPlan", "Floor Plan")
                                        .Replace("CeilingPlan", "Ceiling Plan")
                                        .Replace("ThreeD", "3D View")
                                        .Replace("DrawingSheet", "Sheet");
                                    }
                                    catch { }
                                }

                                string elemDesc = category;
                                if (!string.IsNullOrEmpty(viewType))
                                    elemDesc += " : " + viewType;
                                if (!string.IsNullOrEmpty(name))
                                    elemDesc += " : " + name;
                                elemDesc += " : id " + eid.ToString();

                                FailureMessages.Add(severityTag + " " + elemDesc + " - " + desc);
                            }
                            else
                            {
                                FailureMessages.Add(severityTag + " id " + eid.ToString() + " - " + desc);
                            }
                        }
                        catch
                        {
                            FailureMessages.Add(severityTag + " id " + eid.ToString() + " - " + desc);
                        }
                    }
                }
                else
                {
                    FailureMessages.Add(severityTag + " " + desc);
                }

                // Store for pass 2
                if (severity == FailureSeverity.Warning)
                    warningFailures.Add(failure);
                else if (severity == FailureSeverity.Error)
                    errorFailures.Add(failure);
            }

            // PASS 2: Now resolve/delete after all messages collected
            foreach (var warning in warningFailures)
            {
                try { failuresAccessor.DeleteWarning(warning); } catch { }
            }
            foreach (var error in errorFailures)
            {
                try { failuresAccessor.ResolveFailure(error); } catch { }
            }

            return FailureProcessingResult.Continue;
        }
    }
}