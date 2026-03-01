// AnalyticsSnapshot.cs
// Location: Analytics\AnalyticsSnapshot.cs

using Newtonsoft.Json;
using System;

namespace ScheduleLink.Analytics
{
    /// <summary>
    /// Analytics data stored in local JSON file.
    /// </summary>
    public class AnalyticsSnapshot
    {
        public string AppVersion { get; set; }
        public string RevitVersion { get; set; }
        public DateTime FirstUse { get; set; }
        public DateTime LastUse { get; set; }
        public int TotalSessions { get; set; }
        public int TotalOperations { get; set; }
        public int TotalImports { get; set; }
        public int TotalExports { get; set; }
        public bool HasRated { get; set; }
        public DateTime? LastRatingPrompt { get; set; }
        public DateTime? LastUpdateCheck { get; set; }
        public int RatingPromptCount { get; set; }
        public int NoThanksCount { get; set; }

        [JsonIgnore]
        public int CurrentOperationCount { get; set; } = 0;

        public int DaysSinceFirstUse
        {
            get
            {
                if (FirstUse == DateTime.MinValue)
                    return 0;
                return (int)(DateTime.Now.Date - FirstUse.Date).TotalDays;
            }
        }
    }
}
