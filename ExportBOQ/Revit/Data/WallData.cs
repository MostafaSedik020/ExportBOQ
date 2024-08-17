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
    public class WallData : ElementData
    {
        /// <summary>
        ///Get all the data from shear walls in Revit
        /// <param name="doc">The Revit document to filter.</param>
        /// </summary>
        public static List<WallData> GetWallData(Document doc,string usageParam)
        {
            //filter wall element usage only from other Usages
            var walls = FilterElements.GetWallFilter(doc)
                                         .Where(b => b.LookupParameter("USAGE").AsValueString() == usageParam)
                                         .ToList();

            //Initialze list of beams for the data

            var wallData = new List<WallData>();

            //get the walls volume and convert into cubic meter
            foreach (var wall in walls)
            {
                var volumeParam = wall.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                double volume = volumeParam != null ? volumeParam.AsDouble() : 0;
                double convertedVolume = UnitConverter.convertUnitsToCubicMeters(volume);


                //var areaParam = wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                //double area = areaParam != null ? volumeParam.AsDouble() : 0;
                double area = wall != null ? wall.LookupParameter("Area").AsDouble() : 0;
                double convertedArea = UnitConverter.convertUnitsToSquareMeters(area);

                string type = wall.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString();

                string usage = wall.LookupParameter("USAGE").AsValueString();

                wallData.Add(new WallData
                {
                    Volume = convertedVolume,
                    Type = type,
                    Usage = usage,
                    Area = convertedArea,

                });

            }
            return wallData;
        }
    }
}
