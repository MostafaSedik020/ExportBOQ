using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ExportBOQ.Revit.Utils;

namespace ExportBOQ.Revit.Data
{
    public class MiscData : ElementData
    {
        public static double GetMembraneArea(Document doc)
        {
            List<WallData> rWallData = WallData.GetWallData(doc, "Retaining Wall");
            List<WallData> tankWallData = WallData.GetWallData(doc, "Tank");
            double memArea1 = rWallData.Sum(x => x.Area);
            double memArea2 = tankWallData.Sum(x => x.Area);
            return memArea1+memArea2;

        }

        public static double GetPolySheetArea(Document doc)
        {
            var foundations = FoundationData.GetFoundationData(doc, "RC");
            var floors = FloorData.GetFloorData(doc, "Slab on grade");
            double polyarea1 = foundations.Sum(x => x.Area);
            
            double polyarea2 = floors.Sum(x => x.Area);
            return polyarea1+polyarea2;
        }
        public static double GetBitPaint(Document doc)
        {
            var foundations = FoundationData.GetFoundationData(doc, "RC");
            double bitArea = foundations.Sum(x => x.Area);

            return bitArea;
        }
    }
}
