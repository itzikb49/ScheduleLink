// LicenseManager.cs
// Location: Services\LicenseManager.cs

using Autodesk.Revit.UI;
using System;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;

namespace ScheduleLink.Services
{
    /// <summary>
    /// License management and entitlement checking.
    /// </summary>
    public static class LicenseManager
    {
        // Set to false for testing, true for production
        public const bool ENABLE_LICENSE = true;
        public const bool ENABLE_RATING = true;

        /// <summary>
        /// Checks if user has a valid license via Autodesk Entitlement API.
        /// </summary>
        public static bool IsLicensed(UIApplication uiApp)
        {
            return AutodeskEntitlement.IsUserEntitled(uiApp);
        }
    }

    /// <summary>
    /// Autodesk Entitlement API check.
    /// </summary>
    public static class AutodeskEntitlement
    {
        // TODO: Replace with actual APP_ID from Autodesk App Store after publishing
        private const string APP_ID = "SCHEDULELINK_APP_ID";

        public static bool IsUserEntitled(UIApplication uiApp)
        {
            string userId = uiApp.Application.Username;
            string url = $"https://apps.autodesk.com/webservices/checkentitlement" +
                         $"?userid={userId}&appid={APP_ID}";
            try
            {
                using (var client = new HttpClient())
                {
                    var result = client.GetStringAsync(url).Result;
                    return result.Contains("true");
                }
            }
            catch
            {
                return false;
            }
        }
    }

    /// <summary>
    /// Encryption helper utilities.
    /// </summary>
    public static class EncryptionHelper
    {
        private static readonly byte[] Key = { 73, 66, 45, 66, 73, 77, 50, 48, 50, 54 }; // "IB-BIM2026"

        public static string Encrypt(string text)
        {
            byte[] data = Encoding.UTF8.GetBytes(text);
            byte[] encrypted = new byte[data.Length];
            for (int i = 0; i < data.Length; i++)
                encrypted[i] = (byte)(data[i] ^ Key[i % Key.Length]);
            return Convert.ToBase64String(encrypted);
        }

        public static string Decrypt(string encryptedText)
        {
            byte[] encrypted = Convert.FromBase64String(encryptedText);
            byte[] decrypted = new byte[encrypted.Length];
            for (int i = 0; i < encrypted.Length; i++)
                decrypted[i] = (byte)(encrypted[i] ^ Key[i % Key.Length]);
            return Encoding.UTF8.GetString(decrypted);
        }

        public static string GetHash(string text)
        {
            using (var sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(text + "IB-BIM-SL-SALT"));
                return BitConverter.ToString(bytes).Replace("-", "");
            }
        }
    }
}
