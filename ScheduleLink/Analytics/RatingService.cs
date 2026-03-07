// RatingService.cs
// Location: Analytics\RatingService.cs

using System;
using System.Diagnostics;
using Autodesk.Revit.UI;
using ScheduleLink.Helpers;
using ScheduleLink.Services;
using ScheduleLink.Views;

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
            if (stats.TotalOperations < 25)
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
            // Use styled 3-option dialog
            int choice = MainDialog.ShowRatingDialog(
                "Enjoying ScheduleLink?",
                "Thanks for using ScheduleLink!\n\n" +
                "We'd be grateful if you could rate us on the Autodesk App Store!\n\n" +
                "Your rating helps us improve the tool, reach more BIM professionals, " +
                "and keep developing new features.\n\n" +
                "It takes just a few seconds - thank you!");

            if (choice == 1) // Yes - go to App Store
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = AppStoreReviewUrl,
                        UseShellExecute = true
                    });
                    MainDialog.ShowSuccess("Thank You!",
                        "The App Store is opening in your browser.",
                        "Thank you for taking the time to rate us!");
                }
                catch { }
                AnalyticsService.MarkUserAsRated();
            }
            else if (choice == 2) // Maybe later
            {
                AnalyticsService.MarkRatingPromptShown();
                Logger.Info(Logger.LogCategory.General, "Rating: User selected 'Maybe later'");
            }
            else if (choice == 3) // No thanks
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