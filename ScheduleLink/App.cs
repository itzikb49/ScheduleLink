using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace ScheduleLink
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
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

                // Load embedded icon
                if (btn != null)
                {
                    BitmapSource icon32 = GetEmbeddedImage("ScheduleLink.Resources.Schedule_1_32.png");
                    BitmapSource icon16 = GetEmbeddedImage("ScheduleLink.Resources.Schedule_1_16.png");

                    if (icon32 != null) btn.LargeImage = icon32;
                    if (icon16 != null) btn.Image = icon16;
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("ScheduleLink Error", "Failed to initialize:\n" + ex.Message);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
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
