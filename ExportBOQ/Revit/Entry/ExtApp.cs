using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.IO;
using System.Windows.Media;
using System.Drawing.Imaging;

namespace ExportBOQ.Revit.Entry
{
    public class ExtApp : IExternalApplication
    {
        private UIControlledApplication uicApp;
        public Result OnShutdown(UIControlledApplication uicApp)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication uicApp)
        {
            this.uicApp = uicApp;
            CreatePushButton();

            return Result.Succeeded;
        }

        private void CreatePushButton()
        {
            string tabName = "Trusst";
            string panelName = "Export Data";

            try
            {
                // Create the ribbon tab if it doesn't already exist
                uicApp.CreateRibbonTab(tabName);
            }
            catch (Exception)
            {
                // Tab already exists, continue
            }

            // Get or create the ribbon panel
            RibbonPanel panel = uicApp.GetRibbonPanels(tabName)
                                       .FirstOrDefault(p => p.Name == panelName)
                                       ?? uicApp.CreateRibbonPanel(tabName, panelName);

            // Ensure the panel was created successfully
            if (panel != null)
            {
                // Get the executing assembly path
                Assembly assembly = Assembly.GetExecutingAssembly();
                string assemblyPath = assembly.Location;

                // Create push button data
                PushButtonData pbData = new PushButtonData("ExtApp_btn", "Export BOQ", assemblyPath, typeof(ExtCmd).FullName);

                // Add the push button to the panel
                PushButton pb = panel.AddItem(pbData) as PushButton;
                pb.ToolTip = "Export the Quantity of multiple elements to Excel";

                //pb.LargeImage = new BitmapImage(new Uri($@"{Path.GetDirectoryName(assembly.Location)}\boq32.png"));

                pb.LargeImage = GetImageSource("ExportBOQ.Resources.bill-of-quantities-32.png");
            }
        }
        private ImageSource GetImageSource(string imageFullName)
        {
            Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(imageFullName);
            PngBitmapDecoder decoder = new PngBitmapDecoder(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
            
            return decoder.Frames[0];
        }
    }
}
