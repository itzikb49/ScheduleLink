// Logger.cs - ScheduleLink logging system
using System;
using System.IO;

namespace ScheduleLink.Helpers
{
    /// <summary>
    /// File logger for debugging and error tracking
    /// Main log: ScheduleLink.log - cleared on each Revit start
    /// Export/Import log: ExportImport.log - cleared on each Export/Import operation
    /// </summary>
    public static class Logger
    {
        private static string _logPath;
        private static StreamWriter _logWriter;
        private static readonly object _lock = new object();

        // Export/Import dedicated log
        private static string _exportImportLogPath;
        private static StreamWriter _exportImportLogWriter;
        private static readonly object _exportImportLock = new object();

        public enum LogCategory
        {
            Main,
            General,
            Import,
            Export,
            Validation,
            UI
        }

        #region Main Logger (ScheduleLink.log)

        public static void Initialize()
        {
            Initialize(null, null);
        }

        public static void Initialize(string appName, string version)
        {
            try
            {
                string logFolder = @"C:\ProgramData\IB-BIM\ScheduleLink\Logs";

                if (!Directory.Exists(logFolder))
                {
                    Directory.CreateDirectory(logFolder);
                }

                _logPath = Path.Combine(logFolder, "ScheduleLink.log");
                _exportImportLogPath = Path.Combine(logFolder, "ExportImport.log");

                // append: false = overwrite the file each session
                _logWriter = new StreamWriter(_logPath, append: false);
                _logWriter.AutoFlush = true;

                Info(LogCategory.General, $"=== ScheduleLink Log Started - {DateTime.Now} ===");
                if (!string.IsNullOrEmpty(appName))
                {
                    Info(LogCategory.General, $"Application: {appName} v{version}");
                }
                Info(LogCategory.General, $"Log file: {_logPath}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Logger initialization failed: {ex.Message}");
            }
        }

        public static void Close()
        {
            try
            {
                if (_logWriter != null)
                {
                    Info(LogCategory.General, "=== Log Closed ===");
                    _logWriter.Close();
                    _logWriter.Dispose();
                    _logWriter = null;
                }
            }
            catch { }
        }

        public static void Shutdown()
        {
            Close();
            CloseExportImportLog();
        }

        #endregion

        #region Export/Import Dedicated Logger (ExportImport.log)

        public static void InitializeExportImportLog(string operationName)
        {
            try
            {
                string logFolder = @"C:\ProgramData\IB-BIM\ScheduleLink\Logs";

                if (!Directory.Exists(logFolder))
                {
                    Directory.CreateDirectory(logFolder);
                }

                _exportImportLogPath = Path.Combine(logFolder, "ExportImport.log");

                if (_exportImportLogWriter != null)
                {
                    _exportImportLogWriter.Close();
                    _exportImportLogWriter.Dispose();
                }

                _exportImportLogWriter = new StreamWriter(_exportImportLogPath, append: false);
                _exportImportLogWriter.AutoFlush = true;

                WriteToExportImportLog("==============================================");
                WriteToExportImportLog($"{operationName} - {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                WriteToExportImportLog("==============================================");
            }
            catch (Exception ex)
            {
                Error(LogCategory.Export, "Failed to initialize Export/Import log", ex);
            }
        }

        public static void CompleteExportImportLog(bool success, string summary = null)
        {
            try
            {
                WriteToExportImportLog("----------------------------------------------");
                WriteToExportImportLog($"RESULT: {(success ? "SUCCESS" : "FAILED")}");
                if (!string.IsNullOrEmpty(summary))
                {
                    WriteToExportImportLog($"Summary: {summary}");
                }
                WriteToExportImportLog($"Completed: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                WriteToExportImportLog("==============================================");
            }
            catch { }
            finally
            {
                CloseExportImportLog();
            }
        }

        public static void CloseExportImportLog()
        {
            try
            {
                if (_exportImportLogWriter != null)
                {
                    _exportImportLogWriter.Close();
                    _exportImportLogWriter.Dispose();
                    _exportImportLogWriter = null;
                }
            }
            catch { }
        }

        private static void WriteToExportImportLog(string message)
        {
            if (_exportImportLogWriter == null)
                return;

            try
            {
                lock (_exportImportLock)
                {
                    _exportImportLogWriter.WriteLine(message);
                }
            }
            catch { }
        }

        public static string GetExportImportLogPath()
        {
            return _exportImportLogPath;
        }

        #endregion

        #region Logging Methods

        public static void LogRevitInfo(Autodesk.Revit.ApplicationServices.Application app)
        {
            try
            {
                if (app != null)
                {
                    Info(LogCategory.General, $"Revit Version: {app.VersionName} ({app.VersionBuild})");
                    Info(LogCategory.General, $"Revit Language: {app.Language}");
                }
            }
            catch (Exception ex)
            {
                Warning(LogCategory.General, "Could not log Revit info", ex.Message);
            }
        }

        public static void Info(LogCategory category, string message)
        {
            WriteLog("INFO", category, message);

            if (category == LogCategory.Export || category == LogCategory.Import)
            {
                WriteToExportImportLog($"[INFO] {message}");
            }
        }

        public static void Warning(LogCategory category, string message, string details = null)
        {
            string fullMessage = message + (details != null ? $" - {details}" : "");
            WriteLog("WARN", category, fullMessage);

            if (category == LogCategory.Export || category == LogCategory.Import)
            {
                WriteToExportImportLog($"[WARNING] {fullMessage}");
            }
        }

        public static void Error(LogCategory category, string message, Exception ex = null)
        {
            string errorMsg = message;
            if (ex != null)
            {
                errorMsg += $" | Exception: {ex.GetType().Name} - {ex.Message}";
                if (ex.StackTrace != null)
                {
                    errorMsg += $"\nStack: {ex.StackTrace}";
                }
            }
            WriteLog("ERROR", category, errorMsg);

            if (category == LogCategory.Export || category == LogCategory.Import)
            {
                WriteToExportImportLog($"[ERROR] {message}");
                if (ex != null)
                {
                    WriteToExportImportLog($"[ERROR] Exception: {ex.GetType().Name}");
                    WriteToExportImportLog($"[ERROR] Message: {ex.Message}");
                }
            }
        }

        public static void Debug(LogCategory category, string message)
        {
#if DEBUG
            WriteLog("DEBUG", category, message);
#endif
        }

        private static void WriteLog(string level, LogCategory category, string message)
        {
            if (_logWriter == null)
                return;

            try
            {
                lock (_lock)
                {
                    string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                    string logEntry = $"[{timestamp}] [{level,-5}] [{category,-12}] {message}";
                    _logWriter.WriteLine(logEntry);
                }
            }
            catch { }
        }

        public static string GetLogPath()
        {
            return _logPath;
        }

        #endregion
    }
}
