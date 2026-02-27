// RevitCompatibilityExtensions.cs
// Compatibility layer for Revit 2023-2026+
// Handles ElementId API changes between versions
// Based on TypeManagerPro compatibility layer

using Autodesk.Revit.DB;
using System.Reflection;

namespace ScheduleLink.Helpers
{
    /// <summary>
    /// Extension Methods to support different Revit versions (2023-2026+)
    /// Uses Reflection for full compile-time compatibility
    /// </summary>
    public static class RevitCompatibilityExtensions
    {
        // ════════════════════════════════════════════════════════════════
        // ElementId Value Extraction
        // ════════════════════════════════════════════════════════════════

        /// <summary>
        /// Returns the integer value of ElementId - compatible with all Revit versions
        /// </summary>
        public static int GetIdValue(this ElementId elementId)
        {
            if (elementId == null)
                return -1;

            try
            {
                // Try Value property first (Revit 2024+)
                var valueProperty = typeof(ElementId).GetProperty("Value");
                if (valueProperty != null)
                {
                    var value = valueProperty.GetValue(elementId);
                    return (int)(long)value;
                }
            }
            catch { }

            try
            {
                // Fallback to IntegerValue (Revit 2023)
                var intProperty = typeof(ElementId).GetProperty("IntegerValue");
                if (intProperty != null)
                {
                    return (int)intProperty.GetValue(elementId);
                }
            }
            catch { }

            return -1;
        }

        /// <summary>
        /// Returns the value of ElementId as long
        /// </summary>
        public static long GetIdValueLong(this ElementId elementId)
        {
            if (elementId == null)
                return -1;

            try
            {
                var valueProperty = typeof(ElementId).GetProperty("Value");
                if (valueProperty != null)
                {
                    return (long)valueProperty.GetValue(elementId);
                }
            }
            catch { }

            try
            {
                var intProperty = typeof(ElementId).GetProperty("IntegerValue");
                if (intProperty != null)
                {
                    return (int)intProperty.GetValue(elementId);
                }
            }
            catch { }

            return -1;
        }

        /// <summary>
        /// Checks if ElementId is valid
        /// </summary>
        public static bool IsValidId(this ElementId elementId)
        {
            if (elementId == null)
                return false;

            if (elementId == ElementId.InvalidElementId)
                return false;

            return elementId.GetIdValue() != -1;
        }

        // ════════════════════════════════════════════════════════════════
        // ElementId Creation
        // ════════════════════════════════════════════════════════════════

        /// <summary>
        /// Creates ElementId from int value - compatible with all versions
        /// </summary>
        public static ElementId ToElementId(this int value)
        {
            return ToElementId((long)value);
        }

        /// <summary>
        /// Creates ElementId from long value - compatible with all versions
        /// </summary>
        public static ElementId ToElementId(this long value)
        {
            try
            {
                // Try ElementId(long) constructor first (Revit 2024+)
                var longConstructor = typeof(ElementId).GetConstructor(new[] { typeof(long) });
                if (longConstructor != null)
                {
                    return (ElementId)longConstructor.Invoke(new object[] { value });
                }
            }
            catch { }

            try
            {
                // Fallback to ElementId(int) constructor (Revit 2023)
                var intConstructor = typeof(ElementId).GetConstructor(new[] { typeof(int) });
                if (intConstructor != null)
                {
                    return (ElementId)intConstructor.Invoke(new object[] { (int)value });
                }
            }
            catch { }

            return ElementId.InvalidElementId;
        }

        // ════════════════════════════════════════════════════════════════
        // Revit Version Detection
        // ════════════════════════════════════════════════════════════════

        private static int? _cachedRevitVersion = null;

        /// <summary>
        /// Gets the current Revit version year (e.g., 2023, 2024, 2025)
        /// </summary>
        public static int GetRevitVersion()
        {
            if (_cachedRevitVersion.HasValue)
                return _cachedRevitVersion.Value;

            try
            {
                var assembly = typeof(ElementId).Assembly;
                var version = assembly.GetName().Version;
                int year = 2000 + version.Major;
                _cachedRevitVersion = year;
                return year;
            }
            catch
            {
                _cachedRevitVersion = 2023;
                return 2023;
            }
        }

        /// <summary>
        /// Checks if current Revit version is 2024 or later
        /// </summary>
        public static bool IsRevit2024OrLater()
        {
            return GetRevitVersion() >= 2024;
        }
    }
}
