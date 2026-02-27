using System;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using ScheduleLink.Models;

namespace ScheduleLink.Services
{
    public static class ExcelExportService
    {
        private static readonly Color CLR_INSTANCE  = Color.FromArgb(198, 239, 206);
        private static readonly Color CLR_TYPE      = Color.FromArgb(255, 235, 156);
        private static readonly Color CLR_READONLY  = Color.FromArgb(255, 199, 206);
        private static readonly Color CLR_ELEMENTID = Color.FromArgb(217, 217, 217);
        private static readonly Color CLR_HEADER_BG = Color.FromArgb(68, 114, 196);
        private static readonly Color CLR_HEADER_FG = Color.White;

        public static string Export(ScheduleExportData data, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add(SanitizeSheetName(data.ScheduleName));
                WriteDataSheet(ws, data);

                var instrWs = package.Workbook.Worksheets.Add("Instructions");
                WriteInstructionsSheet(instrWs, data);

                package.SaveAs(new FileInfo(filePath));
            }

            return filePath;
        }

        private static void WriteDataSheet(ExcelWorksheet ws, ScheduleExportData data)
        {
            int totalCols = data.Columns.Count + 1;

            // Header row
            SetHeader(ws, 1, 1, "Element ID");
            for (int c = 0; c < data.Columns.Count; c++)
            {
                string header = !string.IsNullOrEmpty(data.Columns[c].HeaderText)
                    ? data.Columns[c].HeaderText
                    : data.Columns[c].Name;
                SetHeader(ws, 1, c + 2, header);
            }

            // Data rows
            for (int r = 0; r < data.Rows.Count; r++)
            {
                int excelRow = r + 2;
                var row = data.Rows[r];

                // Element ID (grey)
                var idCell = ws.Cells[excelRow, 1];
                idCell.Value = row.ElementId;
                SetCellColor(idCell, CLR_ELEMENTID);

                // Parameters
                for (int c = 0; c < row.Values.Count && c < data.Columns.Count; c++)
                {
                    int excelCol = c + 2;
                    var cell = ws.Cells[excelRow, excelCol];
                    var colInfo = data.Columns[c];

                    SetTypedValue(cell, row.Values[c], colInfo.StorageType);

                    if (colInfo.IsReadOnly || colInfo.IsCalculated)
                        SetCellColor(cell, CLR_READONLY);
                    else if (colInfo.IsTypeParam)
                        SetCellColor(cell, CLR_TYPE);
                    else
                        SetCellColor(cell, CLR_INSTANCE);
                }
            }

            // --- Borders around all cells ---
            if (data.Rows.Count > 0)
            {
                var dataRange = ws.Cells[1, 1, data.Rows.Count + 1, totalCols];
                dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                dataRange.Style.Border.Top.Color.SetColor(Color.FromArgb(180, 180, 180));
                dataRange.Style.Border.Bottom.Color.SetColor(Color.FromArgb(180, 180, 180));
                dataRange.Style.Border.Left.Color.SetColor(Color.FromArgb(180, 180, 180));
                dataRange.Style.Border.Right.Color.SetColor(Color.FromArgb(180, 180, 180));
            }

            // --- Auto-fit columns based on content ---
            for (int c = 1; c <= totalCols; c++)
            {
                ws.Column(c).AutoFit();
                // Ensure minimum and maximum widths
                if (ws.Column(c).Width < 8)
                    ws.Column(c).Width = 8;
                if (ws.Column(c).Width > 60)
                    ws.Column(c).Width = 60;
            }

            // Freeze header row + ID column
            ws.View.FreezePanes(2, 2);

            // Auto-filter
            if (data.Rows.Count > 0)
                ws.Cells[1, 1, 1, totalCols].AutoFilter = true;
        }

        private static void WriteInstructionsSheet(ExcelWorksheet ws, ScheduleExportData data)
        {
            int r = 1;
            ws.Cells[r, 1].Value = "ScheduleLink - IB-BIM Tools";
            ws.Cells[r, 1].Style.Font.Bold = true;
            ws.Cells[r, 1].Style.Font.Size = 16;

            r += 2;
            WriteLabel(ws, r, "Schedule:", data.ScheduleName); r++;
            WriteLabel(ws, r, "Schedule ID:", data.ScheduleId.ToString()); r++;
            WriteLabel(ws, r, "Rows:", data.Rows.Count.ToString()); r++;
            WriteLabel(ws, r, "Export Date:", DateTime.Now.ToString("yyyy-MM-dd HH:mm")); r++;

            r++;
            ws.Cells[r, 1].Value = "Color Legend:";
            ws.Cells[r, 1].Style.Font.Bold = true;
            ws.Cells[r, 1].Style.Font.Size = 13;
            r++;

            WriteColorLegend(ws, r, CLR_INSTANCE, "  Instance Parameter (editable)"); r++;
            WriteColorLegend(ws, r, CLR_TYPE, "  Type Parameter (shared value)"); r++;
            WriteColorLegend(ws, r, CLR_READONLY, "  Read-Only (DO NOT EDIT)"); r++;
            WriteColorLegend(ws, r, CLR_ELEMENTID, "  Element ID (DO NOT EDIT)"); r++;

            r++;
            ws.Cells[r, 1].Value = "Rules:";
            ws.Cells[r, 1].Style.Font.Bold = true;
            ws.Cells[r, 1].Style.Font.Size = 13;
            r++;

            ws.Cells[r, 1].Value = "1. Do NOT change Element ID column"; r++;
            ws.Cells[r, 1].Value = "2. Do NOT add or remove rows"; r++;
            ws.Cells[r, 1].Value = "3. Do NOT rename column headers"; r++;
            ws.Cells[r, 1].Value = "4. Edit only GREEN and YELLOW cells"; r++;
            ws.Cells[r, 1].Value = "5. Undo in Revit: Ctrl+Z"; r++;

            ws.Column(1).AutoFit(20, 80);
            ws.Column(2).AutoFit(15, 40);
        }

        #region Helpers

        private static void SetHeader(ExcelWorksheet ws, int row, int col, string text)
        {
            var cell = ws.Cells[row, col];
            cell.Value = text;
            cell.Style.Font.Bold = true;
            cell.Style.Font.Color.SetColor(CLR_HEADER_FG);
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(CLR_HEADER_BG);
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            cell.Style.Border.Bottom.Color.SetColor(Color.FromArgb(40, 70, 140));
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        private static void SetCellColor(ExcelRange cell, Color color)
        {
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(color);
        }

        private static void SetTypedValue(ExcelRange cell, string value, string storageType)
        {
            if (string.IsNullOrEmpty(value)) { cell.Value = ""; return; }

            if (storageType == "Integer" && int.TryParse(value, out int i))
            { cell.Value = i; return; }

            if (storageType == "Double")
            {
                if (double.TryParse(value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out double d))
                { cell.Value = d; return; }
                if (double.TryParse(value, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.CurrentCulture, out double d2))
                { cell.Value = d2; return; }
            }

            cell.Value = value;
        }

        private static void WriteLabel(ExcelWorksheet ws, int row, string label, string value)
        {
            ws.Cells[row, 1].Value = label;
            ws.Cells[row, 1].Style.Font.Bold = true;
            ws.Cells[row, 2].Value = value;
        }

        private static void WriteColorLegend(ExcelWorksheet ws, int row, Color color, string text)
        {
            var cell = ws.Cells[row, 1];
            cell.Value = text;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(color);
        }

        private static string SanitizeSheetName(string name)
        {
            string s = name;
            foreach (char c in new[] { '\\', '/', '*', '[', ']', ':', '?' })
                s = s.Replace(c, '_');
            return s.Length > 31 ? s.Substring(0, 31) : s;
        }

        #endregion
    }
}
