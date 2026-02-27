using System;
using System.IO;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ScheduleLink.Models;
using ScheduleLink.Services;
using ScheduleLink.Views;

namespace ScheduleLink.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ScheduleLinkCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            try
            {
                var schedules = ScheduleReaderService.GetAllSchedules(doc);

                if (schedules.Count == 0)
                {
                    TaskDialog.Show("ScheduleLink", "No schedules found in the project.");
                    return Result.Failed;
                }

                var dialog = new MainDialog(schedules, doc);
                bool? dialogResult = dialog.ShowDialog();

                if (dialogResult != true)
                    return Result.Cancelled;

                if (dialog.IsExportMode)
                    return ExecuteExport(doc, dialog.SelectedSchedule);
                else
                    return ExecuteImport(doc);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("ScheduleLink - Error", ex.Message);
                return Result.Failed;
            }
        }

        private Result ExecuteExport(Document doc, ViewSchedule schedule)
        {
            // Read data inside a transaction (some Revit versions require it)
            ScheduleExportData data;
            using (var tx = new Transaction(doc, "ScheduleLink Export Read"))
            {
                tx.Start();
                data = ScheduleReaderService.ExtractScheduleData(doc, schedule);
                tx.RollBack(); // No changes, just reading
            }

            if (data.Rows.Count == 0)
            {
                TaskDialog.Show("ScheduleLink", "Schedule contains no data rows.");
                return Result.Failed;
            }

            var saveDlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save Schedule to Excel",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = SanitizeFileName(schedule.Name) + ".xlsx",
                DefaultExt = ".xlsx"
            };

            if (saveDlg.ShowDialog() != true)
                return Result.Cancelled;

            string filePath = ExcelExportService.Export(data, saveDlg.FileName);

            var td = new TaskDialog("ScheduleLink - Export Complete");
            td.MainInstruction = "Export completed successfully!";
            td.MainContent = "Schedule: " + data.ScheduleName + "\n"
                           + "Rows: " + data.Rows.Count + "\n"
                           + "Columns: " + data.Columns.Count + "\n\n"
                           + filePath;
            td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
            td.DefaultButton = TaskDialogResult.Yes;
            td.FooterText = "Open the file now?";

            if (td.Show() == TaskDialogResult.Yes)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }

            return Result.Succeeded;
        }

        private Result ExecuteImport(Document doc)
        {
            var openDlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select Excel File to Import",
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                FilterIndex = 1
            };

            if (openDlg.ShowDialog() != true)
                return Result.Cancelled;

            ScheduleExportData excelData = ExcelImportService.ReadExcelFile(openDlg.FileName);
            if (excelData == null || excelData.Rows.Count == 0)
            {
                TaskDialog.Show("ScheduleLink", "Could not read data from the Excel file.");
                return Result.Failed;
            }

            int editable = 0, readOnly = 0;
            foreach (var col in excelData.Columns)
            {
                if (col.IsReadOnly) readOnly++;
                else editable++;
            }

            var confirmDlg = new TaskDialog("ScheduleLink - Confirm Import");
            confirmDlg.MainInstruction = "Import data to Revit?";
            confirmDlg.MainContent = "Schedule: " + excelData.ScheduleName + "\n"
                                   + "Rows: " + excelData.Rows.Count + "\n"
                                   + "Editable columns: " + editable + "\n"
                                   + "Read-only columns (skip): " + readOnly;
            confirmDlg.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
            confirmDlg.DefaultButton = TaskDialogResult.Yes;
            confirmDlg.FooterText = "You can undo with Ctrl+Z";

            if (confirmDlg.Show() != TaskDialogResult.Yes)
                return Result.Cancelled;

            // Single transaction for import (supports undo)
            ImportResult result = ExcelImportService.ImportToRevit(doc, excelData);

            var sb = new StringBuilder();
            sb.AppendLine("Rows: " + result.TotalRows);
            sb.AppendLine("Updated: " + result.UpdatedParams);
            sb.AppendLine("Unchanged: " + result.SkippedUnchanged);
            sb.AppendLine("Read-only (skipped): " + result.SkippedReadOnly);
            sb.AppendLine("Not found: " + result.ParamNotFound);
            sb.AppendLine("Failed: " + result.FailedParams);

            if (result.ElementsNotFound > 0)
                sb.AppendLine("Elements not found: " + result.ElementsNotFound);

            if (result.Errors.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Errors:");
                for (int i = 0; i < Math.Min(result.Errors.Count, 5); i++)
                    sb.AppendLine("  " + result.Errors[i]);
            }

            string title = result.UpdatedParams > 0 ? "Import Complete" : "No Changes";
            TaskDialog.Show("ScheduleLink - " + title, sb.ToString());

            return Result.Succeeded;
        }

        private string SanitizeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }
    }
}
