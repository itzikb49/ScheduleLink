// AnalyticsService.cs
// Location: Analytics\AnalyticsService.cs

using Newtonsoft.Json;
using System;
using System.IO;
using Autodesk.Revit.DB;
using ScheduleLink.Helpers;

namespace ScheduleLink.Analytics
{
    /// <summary>
    /// Manages local analytics data (JSON).
    /// </summary>
    public static class AnalyticsService
    {
        private static readonly string FolderPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "IB-BIM", "ScheduleLink");

        private static readonly string FilePath =
            Path.Combine(FolderPath, "analytics.json");

        private static AnalyticsSnapshot _cache;

        public static AnalyticsSnapshot GetSnapshot()
        {
            if (_cache != null)
                return _cache;

            try
            {
                if (!Directory.Exists(FolderPath))
                    Directory.CreateDirectory(FolderPath);

                if (!File.Exists(FilePath))
                {
                    _cache = CreateNewSnapshot();
                    SaveSnapshot();
                }
                else
                {
                    var json = File.ReadAllText(FilePath);
                    _cache = JsonConvert.DeserializeObject<AnalyticsSnapshot>(json)
                             ?? CreateNewSnapshot();
                }
            }
            catch
            {
                _cache = CreateNewSnapshot();
            }

            return _cache;
        }

        private static AnalyticsSnapshot CreateNewSnapshot()
        {
            var now = DateTime.Now;
            return new AnalyticsSnapshot
            {
                FirstUse = now,
                LastUse = now,
                TotalSessions = 0,
                TotalOperations = 0,
                TotalImports = 0,
                TotalExports = 0,
                HasRated = false,
                LastRatingPrompt = null,
                RatingPromptCount = 0
            };
        }

        public static void SaveSnapshot()
        {
            try
            {
                if (!Directory.Exists(FolderPath))
                    Directory.CreateDirectory(FolderPath);

                string json = JsonConvert.SerializeObject(_cache, Formatting.Indented);
                File.WriteAllText(FilePath, json);
            }
            catch { }
        }

        public static void TrackSessionStarted(Document doc)
        {
            var s = GetSnapshot();
            s.LastUse = DateTime.Now;
            s.TotalSessions++;

            try
            {
                s.RevitVersion = doc.Application.VersionNumber + " (" + doc.Application.VersionName + ")";
            }
            catch
            {
                s.RevitVersion = "Unknown";
            }

            if (string.IsNullOrEmpty(s.AppVersion))
                s.AppVersion = "1.0.0";

            SaveSnapshot();
        }

        public static void TrackOperation(int imports = 0, int exports = 0)
        {
            var s = GetSnapshot();
            int totalDelta = imports + exports;

            if (totalDelta == 0)
            {
                s.LastUse = DateTime.Now;
                SaveSnapshot();
                return;
            }

            if (s.TotalOperations >= 1_000_000)
            {
                SaveSnapshot();
                return;
            }

            s.LastUse = DateTime.Now;
            s.TotalOperations += totalDelta;
            s.TotalImports += imports;
            s.TotalExports += exports;

            SaveSnapshot();
        }

        public static void MarkUserAsRated()
        {
            var s = GetSnapshot();
            s.HasRated = true;
            s.LastRatingPrompt = DateTime.Now;
            s.RatingPromptCount++;
            SaveSnapshot();
        }

        public static void MarkRatingPromptShown()
        {
            var s = GetSnapshot();
            s.LastRatingPrompt = DateTime.Now;
            s.RatingPromptCount++;
            SaveSnapshot();
        }

        public static void IncrementNoThanksCount()
        {
            var s = GetSnapshot();
            s.NoThanksCount++;
            s.LastRatingPrompt = DateTime.Now;
            s.RatingPromptCount++;
            SaveSnapshot();
        }
    }
}
