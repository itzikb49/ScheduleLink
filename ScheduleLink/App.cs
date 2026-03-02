// App.cs

using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;
using ScheduleLink.Analytics;
using ScheduleLink.Helpers;

namespace ScheduleLink
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Initialize Logger (overwrites previous session log)
                Logger.Initialize("ScheduleLink", "1.0.0");
                Logger.Info(Logger.LogCategory.General, "Revit: " + application.ControlledApplication.VersionName);
                Logger.Info(Logger.LogCategory.General, "Assembly: " + Assembly.GetExecutingAssembly().Location);

                string tabName = "IB-BIM Tools";
                try { application.CreateRibbonTab(tabName); }
                catch { }

                RibbonPanel panel = application.CreateRibbonPanel(tabName, "ScheduleLink");
                string assemblyPath = Assembly.GetExecutingAssembly().Location;

                var btnData = new PushButtonData(
                    "ScheduleLink",
                    "Import/Export\nSchedule",
                    assemblyPath,
                    "ScheduleLink.Commands.ScheduleLinkCommand")
                {
                    ToolTip = "Export/Import schedules to/from Excel",
                    LongDescription =
                        "Export any Revit schedule to Excel and import edited data back.\n" +
                        "Color-coded columns: Green=Editable, Yellow=Type, Red=ReadOnly.\n" +
                        "Supports Ctrl+Z undo."
                };

                PushButton btn = panel.AddItem(btnData) as PushButton;

                if (btn != null)
                {
                    BitmapSource icon32 = GetEmbeddedImage("ScheduleLink.Resources.Schedule_1_32.png");
                    BitmapSource icon16 = GetEmbeddedImage("ScheduleLink.Resources.Schedule_1_16.png");

                    if (icon32 != null) btn.LargeImage = icon32;
                    if (icon16 != null) btn.Image = icon16;
                    Logger.Info(Logger.LogCategory.General, "Icons loaded: 32=" + (icon32 != null) + " 16=" + (icon16 != null));

                    // F1 help - opens online documentation
                    btn.SetContextualHelp(new ContextualHelp(
                        ContextualHelpType.Url,
                        "https://itzikb49.github.io/IB-BIM-ScheduleLink-Docs/UserGuide"));
                }

                Logger.Info(Logger.LogCategory.General, "ScheduleLink initialized successfully");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Logger.Error(Logger.LogCategory.General, "Failed to initialize ScheduleLink", ex);
                TaskDialog.Show("ScheduleLink Error", "Failed to initialize:\n" + ex.Message);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            Logger.Shutdown();
            return Result.Succeeded;
        }

        private static BitmapSource GetEmbeddedImage(string name)
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                Stream stream = assembly.GetManifestResourceStream(name);
                if (stream != null)
                {
                    BitmapSource image = BitmapFrame.Create(stream);
                    return image;
                }
            }
            catch { }
            return null;
        }
    }
}
