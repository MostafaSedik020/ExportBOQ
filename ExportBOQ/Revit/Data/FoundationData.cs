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
    public class FoundationData : ElementData
    {
        /// <summary>
        ///Get all the data from Foundation in Revit
        /// </summary>
        public static List<FoundationData> GetFoundationData(Document doc , string matType)
        {
            //filter columns element 
            var foundations = FilterElements.GetFoundationFilter(doc);


            //Initialze list of stair for the data

            var foundationData = new List<FoundationData>();

            ElementId materialID = null;
            if (matType == "RC")
            {
                Material material = FilterElements.GetRcMat(doc);
                if (material != null)
                {
                   materialID = material.Id;
                }
            }
            else if (matType == "PC")
            {
                Material material = FilterElements.GetPcMat(doc);
                if (material != null)
                {
                    materialID = material.Id;
                }
            }


            //get the foundation volume and convert into cubic meter
            foreach (var foundation in foundations)
            {
                string type = foundation.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString();

                double? volume = foundation.GetMaterialVolume(materialID);
                double convertedVolume = volume != null ? UnitConverter.convertUnitsToCubicMeters((double)volume) : 0;

                double convertedArea = 0;

                if (!type.Contains("PC"))
                {
                    var areaParam = foundation.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                    double area = areaParam != null ? areaParam.AsDouble() : 0;
                     convertedArea = UnitConverter.convertUnitsToSquareMeters(area);
                }
                

                foundationData.Add(new FoundationData
                {
                    Volume = convertedVolume,
                    Type = type,
                    Area = convertedArea,
                });
            }
            return foundationData;
        }
    }
}
