using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ExportBOQ.Excel;
using ExportBOQ.Revit.Data;
using ExportBOQ.UI;


namespace ExportBOQ.Revit.Entry
{
    [Transaction(TransactionMode.Manual)]
    public class ExtCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            Document doc = uidoc.Document;

            MainWindow mainWindow = new MainWindow(doc);
            mainWindow.ShowDialog();



            return Result.Succeeded;
        }
    }
}
