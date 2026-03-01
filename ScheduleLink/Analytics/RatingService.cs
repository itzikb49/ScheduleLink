// RatingService.cs
// Location: Analytics\RatingService.cs

using System;
using System.Diagnostics;
using Autodesk.Revit.UI;
using ScheduleLink.Helpers;
using ScheduleLink.Services;

namespace ScheduleLink.Analytics
{
    /// <summary>
    /// Decides when to ask user for App Store rating.
    /// </summary>
    public static class RatingService
    {
        // TODO: Replace with actual ScheduleLink App Store URL after publishing
        private const string AppStoreReviewUrl =
            "https://apps.autodesk.com/RVT/en/Detail/Index?id=SCHEDULELINK_APP_ID#reviews";

        public static void MaybeAskForRating(UIApplication uiApp)
        {
            if (!LicenseManager.ENABLE_RATING)
                return;

            var stats = AnalyticsService.GetSnapshot();

            if (stats.HasRated)
                return;

            // Condition 1: at least 14 days
            if (stats.DaysSinceFirstUse < 14)
                return;

            // Condition 2: at least 50 operations (Import + Export)
            if (stats.TotalOperations < 50)
                return;

            if (stats.LastRatingPrompt.HasValue)
            {
                var daysSince = (DateTime.Now - stats.LastRatingPrompt.Value).TotalDays;

                if (stats.NoThanksCount > 0)
                {
                    if (stats.NoThanksCount >= 2)
                    {
                        AnalyticsService.MarkUserAsRated();
                        return;
                    }
                    if (daysSince < 14)
                        return;
                }
                else
                {
                    if (daysSince < 2)
                        return;
                }
            }

            ShowRatingDialog();
        }

        private static void ShowRatingDialog()
        {
            var dialog = new TaskDialog("Enjoying ScheduleLink?")
            {
                MainIcon = TaskDialogIcon.TaskDialogIconNone,
                MainInstruction = "Quick question...",
                MainContent =
                    "Thanks for using ScheduleLink!\n\n" +
                    "We'd be grateful if you could rate us on the Autodesk App Store!\n\n" +
                    "Your rating helps us:\n" +
                    "   - Improve the tool for you\n" +
                    "   - Reach more BIM professionals\n" +
                    "   - Keep developing new features\n\n" +
                    "It takes just a few seconds - thank you!",
                CommonButtons = TaskDialogCommonButtons.None
            };

            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                "Yes, take me to the App Store", "Opens your browser (30 seconds)");
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                "Maybe later", "We'll ask again in 2 days");
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3,
                "No thanks", "We'll ask once more in 2 weeks");

            var result = dialog.Show();

            if (result == TaskDialogResult.CommandLink1)
            {
                try
                {
                    Process.Start(AppStoreReviewUrl);
                    TaskDialog.Show("Thank You!",
                        "The App Store is opening in your browser.\n\n" +
                        "Thank you for taking the time to rate us!");
                }
                catch { }
                AnalyticsService.MarkUserAsRated();
            }
            else if (result == TaskDialogResult.CommandLink2)
            {
                AnalyticsService.MarkRatingPromptShown();
                Logger.Info(Logger.LogCategory.General, "Rating: User selected 'Maybe later'");
            }
            else if (result == TaskDialogResult.CommandLink3)
            {
                AnalyticsService.IncrementNoThanksCount();
                var stats = AnalyticsService.GetSnapshot();
                if (stats.NoThanksCount >= 2)
                {
                    AnalyticsService.MarkUserAsRated();
                    Logger.Info(Logger.LogCategory.General, "Rating: User declined 2 times - won't ask again");
                }
                else
                {
                    Logger.Info(Logger.LogCategory.General, $"Rating: User declined (count: {stats.NoThanksCount})");
                }
            }
        }
    }
}
