// ScheduleLinkCommand.cs
using System;
using System.IO;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ScheduleLink.Analytics;
using ScheduleLink.Helpers;
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
                Logger.Info(Logger.LogCategory.General, "--- Command Execute ---");
                Logger.Info(Logger.LogCategory.General, "Document: " + doc.Title);

                // Track session for analytics
                AnalyticsService.TrackSessionStarted(doc);

                // Check for updates (background, non-blocking)
                UpdateCheckService.CheckForUpdates();

                var schedules = ScheduleReaderService.GetAllSchedules(doc);
                Logger.Info(Logger.LogCategory.General, "Schedules found: " + schedules.Count);

                if (schedules.Count == 0)
                {
                    MainDialog.ShowInfo("ScheduleLink", "No schedules found in the project.");
                    return Result.Failed;
                }

                var dialog = new MainDialog(schedules, doc);
                bool? dialogResult = dialog.ShowDialog();

                if (dialogResult != true)
                    return Result.Cancelled;

                Result result;
                if (dialog.IsExportMode)
                    result = ExecuteExport(doc, dialog.SelectedSchedule);
                else
                    result = ExecuteImport(doc);

                // Ask for rating after successful operation
                if (result == Result.Succeeded)
                {
                    try { RatingService.MaybeAskForRating(uiApp); }
                    catch { }
                }

                return result;
            }
            catch (Exception ex)
            {
                Logger.Error(Logger.LogCategory.General, "Command.Execute", ex);
                message = ex.Message;
                MainDialog.ShowError("ScheduleLink - Error", ex.Message);
                return Result.Failed;
            }
        }

        private Result ExecuteExport(Document doc, ViewSchedule schedule)
        {
            Logger.Info(Logger.LogCategory.Export, "Export: Schedule='" + schedule.Name + "'");
            Logger.InitializeExportImportLog("EXPORT - " + schedule.Name);

            ScheduleExportData data;
            using (var tx = new Transaction(doc, "ScheduleLink Export Read"))
            {
                tx.Start();
                data = ScheduleReaderService.ExtractScheduleData(doc, schedule);
                tx.RollBack();
            }

            Logger.Info(Logger.LogCategory.Export, "Export: Rows=" + data.Rows.Count + " Columns=" + data.Columns.Count);

            if (data.Rows.Count == 0)
            {
                Logger.CompleteExportImportLog(false, "No data rows");
                MainDialog.ShowInfo("ScheduleLink", "Schedule contains no data rows.");
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
            Logger.Info(Logger.LogCategory.Export, "Export: Saved to " + filePath);

            // Track analytics
            AnalyticsService.TrackOperation(exports: 1);

            Logger.CompleteExportImportLog(true,
                "Rows=" + data.Rows.Count + " Columns=" + data.Columns.Count);

            // Styled export complete dialog
            bool openFile = MainDialog.ShowExportComplete(
                data.ScheduleName, data.Rows.Count, data.Columns.Count, filePath);

            if (openFile)
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
            Logger.Info(Logger.LogCategory.Import, "Import: Started");
            Logger.InitializeExportImportLog("IMPORT");

            var openDlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select Excel File to Import",
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                FilterIndex = 1
            };

            if (openDlg.ShowDialog() != true)
                return Result.Cancelled;

            Logger.Info(Logger.LogCategory.Import, "Import: File=" + openDlg.FileName);

            ScheduleExportData excelData = ExcelImportService.ReadExcelFile(openDlg.FileName);
            if (excelData == null || excelData.Rows.Count == 0)
            {
                Logger.Info(Logger.LogCategory.Import, "Import: Failed to read Excel data");
                Logger.CompleteExportImportLog(false, "Failed to read Excel data");
                MainDialog.ShowError("ScheduleLink", "Could not read data from the Excel file.");
                return Result.Failed;
            }

            Logger.Info(Logger.LogCategory.Import, "Import: Rows=" + excelData.Rows.Count + " Columns=" + excelData.Columns.Count);

            int editable = 0, readOnly = 0;
            foreach (var col in excelData.Columns)
            {
                if (col.IsReadOnly) readOnly++;
                else editable++;
            }

            // Styled confirm dialog
            string confirmMsg = "Schedule: " + excelData.ScheduleName + "\n"
                              + "Rows: " + excelData.Rows.Count + "\n"
                              + "Editable columns: " + editable + "\n"
                              + "Read-only columns (skip): " + readOnly;

            bool confirmed = MainDialog.ShowConfirm(
                "Confirm Import",
                confirmMsg,
                "You can undo with Ctrl+Z",
                "Import", "Cancel");

            if (!confirmed)
                return Result.Cancelled;

            ImportResult result = ExcelImportService.ImportToRevit(doc, excelData);

            Logger.Info(Logger.LogCategory.Import, "Import: Updated=" + result.UpdatedParams +
                " Unchanged=" + result.SkippedUnchanged +
                " ReadOnly=" + result.SkippedReadOnly +
                " Failed=" + result.FailedParams +
                " NotFound=" + result.ElementsNotFound);

            // Track analytics
            AnalyticsService.TrackOperation(imports: 1);

            Logger.CompleteExportImportLog(result.UpdatedParams > 0,
                "Updated=" + result.UpdatedParams + " Failed=" + result.FailedParams);

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
                sb.AppendLine("--- Errors (" + result.Errors.Count + ") ---");
                for (int i = 0; i < result.Errors.Count; i++)
                    sb.AppendLine("  " + result.Errors[i]);
            }

            string title = result.FailedParams == 0 && result.UpdatedParams > 0
                ? "Import Complete"
                : result.UpdatedParams > 0
                    ? "Import Complete (with errors)"
                    : "No Changes";

            MainDialog.ShowImportResult(title, sb.ToString(), result.UpdatedParams > 0, result.FailedParams);

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