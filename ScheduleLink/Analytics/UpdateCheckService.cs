// UpdateCheckService.cs
// Location: Analytics\UpdateCheckService.cs

using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Threading.Tasks;
using ScheduleLink.Helpers;
using ScheduleLink.Views;

namespace ScheduleLink.Analytics
{
    /// <summary>
    /// Checks if a newer version is available.
    /// </summary>
    public static class UpdateCheckService
    {
        private const string CURRENT_VERSION = "1.0.0";

        // TODO: Create public repo ScheduleLink-Updates with version.json
        private const string VERSION_CHECK_URL =
            "https://raw.githubusercontent.com/itzikb49/ScheduleLink-Updates/main/version.json";

        private const int DAYS_BETWEEN_CHECKS = 7;

        public static void CheckForUpdates()
        {
            try
            {
                var stats = AnalyticsService.GetSnapshot();
                if (stats == null)
                    return;

                if (!ShouldCheckNow(stats))
                    return;

                Task.Run(() => CheckForUpdatesAsync(stats));
                Logger.Info(Logger.LogCategory.General, "Update check started in background");
            }
            catch (Exception ex)
            {
                Logger.Error(Logger.LogCategory.General, "Update check failed to start", ex);
            }
        }

        private static bool ShouldCheckNow(AnalyticsSnapshot stats)
        {
            try
            {
                if (stats.TotalSessions <= 1)
                    return false;

                if (stats.LastUpdateCheck.HasValue)
                {
                    var daysSinceLastCheck = (DateTime.Now - stats.LastUpdateCheck.Value).TotalDays;
                    if (daysSinceLastCheck < DAYS_BETWEEN_CHECKS)
                        return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static async Task CheckForUpdatesAsync(AnalyticsSnapshot stats)
        {
            HttpClient client = null;

            try
            {
                client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(5);

                string json = await client.GetStringAsync(VERSION_CHECK_URL);

                if (string.IsNullOrWhiteSpace(json))
                    return;

                var versionInfo = JsonConvert.DeserializeObject<VersionInfo>(json);

                if (versionInfo == null || string.IsNullOrEmpty(versionInfo.LatestVersion))
                    return;

                Logger.Info(Logger.LogCategory.General,
                    $"Latest version: {versionInfo.LatestVersion}, Current: {CURRENT_VERSION}");

                stats.LastUpdateCheck = DateTime.Now;
                AnalyticsService.SaveSnapshot();

                if (IsNewerVersion(versionInfo.LatestVersion, CURRENT_VERSION))
                {
                    Logger.Info(Logger.LogCategory.General, "Update available!");
                    ShowUpdateNotification(versionInfo);
                }
            }
            catch (HttpRequestException ex)
            {
                Logger.Info(Logger.LogCategory.General, $"Update check: No internet ({ex.Message})");
            }
            catch (TaskCanceledException)
            {
                Logger.Info(Logger.LogCategory.General, "Update check: Timeout");
            }
            catch (Exception ex)
            {
                Logger.Error(Logger.LogCategory.General, "Update check failed", ex);
            }
            finally
            {
                client?.Dispose();
            }
        }

        private static bool IsNewerVersion(string newVersion, string currentVersion)
        {
            try
            {
                if (string.IsNullOrEmpty(newVersion) || string.IsNullOrEmpty(currentVersion))
                    return false;

                var newParts = newVersion.Split('.');
                var currentParts = currentVersion.Split('.');

                for (int i = 0; i < Math.Min(newParts.Length, currentParts.Length); i++)
                {
                    if (!int.TryParse(newParts[i], out int newNum)) return false;
                    if (!int.TryParse(currentParts[i], out int currentNum)) return false;

                    if (newNum > currentNum) return true;
                    if (newNum < currentNum) return false;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private static void ShowUpdateNotification(VersionInfo versionInfo)
        {
            try
            {
                var isCritical = versionInfo.CriticalUpdate;

                var changelogText = (versionInfo.Changelog != null && versionInfo.Changelog.Length > 0)
                    ? string.Join("\n", versionInfo.Changelog)
                    : "Bug fixes and improvements";

                string title = isCritical
                    ? "Critical Update Available"
                    : "Update Available";

                string message = $"Version {versionInfo.LatestVersion} is now available!\n" +
                    $"You're currently using version {CURRENT_VERSION}.\n\n" +
                    (isCritical ? "This is a critical update with important bug fixes.\n\n" : "") +
                    "What's new:\n" + changelogText;

                bool download;
                if (isCritical)
                {
                    download = MainDialog.ShowWarning(title, message, "Would you like to download the update?");
                }
                else
                {
                    download = MainDialog.ShowConfirm(title, message,
                        "Would you like to download the update?",
                        "Download", "Not Now");
                }

                if (download)
                {
                    try
                    {
                        var url = !string.IsNullOrEmpty(versionInfo.DownloadUrl)
                            ? versionInfo.DownloadUrl
                            : "https://apps.autodesk.com";

                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = url,
                            UseShellExecute = true
                        });
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Logger.Error(Logger.LogCategory.General, "Failed to show update dialog", ex);
            }
        }

        public static string GetCurrentVersion()
        {
            return CURRENT_VERSION;
        }
    }

    internal class VersionInfo
    {
        [JsonProperty("latestVersion")]
        public string LatestVersion { get; set; }

        [JsonProperty("releaseDate")]
        public string ReleaseDate { get; set; }

        [JsonProperty("downloadUrl")]
        public string DownloadUrl { get; set; }

        [JsonProperty("changelog")]
        public string[] Changelog { get; set; }

        [JsonProperty("criticalUpdate")]
        public bool CriticalUpdate { get; set; }

        [JsonProperty("minimumVersion")]
        public string MinimumVersion { get; set; }
    }
}