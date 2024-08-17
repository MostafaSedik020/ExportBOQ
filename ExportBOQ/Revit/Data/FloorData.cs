using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ExportBOQ.Revit.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportBOQ.Revit.Data
{
    public class FloorData : ElementData
    {
        /// <summary>
        ///Get all the data from floor in Revit
        /// </summary>
        public static List<FloorData> GetFloorData(Document doc,string usageParam)
        {
            //filter floor element 
            var floors = FilterElements.GetFloorFilter(doc)
                                       .Where(b => b.LookupParameter("USAGE").AsValueString() == usageParam)
                                       .ToList();

            //Initialze list of floor for the data

            var floorData = new List<FloorData>();

            //get the floor volume and convert into cubic meter
            foreach (var floor in floors)
            {
                var volumeParam = floor.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                double volume = volumeParam != null ? volumeParam.AsDouble() : 0;
                double convetedVolume = UnitConverter.convertUnitsToCubicMeters(volume);

                var areaParam = floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                double area = areaParam != null ? areaParam.AsDouble() : 0;
                double convertedArea = UnitConverter.convertUnitsToSquareMeters(area);

                string type = floor.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString();

                string usage = floor.LookupParameter("USAGE").AsValueString();

                string level = floor.LookupParameter("Level").AsValueString();

                floorData.Add(new FloorData
                {
                    Volume = convetedVolume,
                    Type = type,
                    Usage = usage,
                    Level = level,
                    Area = convertedArea,
                });
            }

            return floorData;
        }
    }
}
