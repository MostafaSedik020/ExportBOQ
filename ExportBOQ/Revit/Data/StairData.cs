using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using ExportBOQ.Revit.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportBOQ.Revit.Data
{
    public class StairData : ElementData
    {
        /// <summary>
        ///Get all the data from Stairs in Revit
        /// </summary>
        public static List<StairData> GetStairData(Document doc)
        {
            //filter columns element 
            var stairs = FilterElements.GetStairFilter(doc);
            

            //Initialze list of stair for the data

            var stairData = new List<StairData>();



            //get the stair volume and convert into cubic meter
            foreach (var stair in stairs)
            {

                var matStair = stair.GetMaterialIds(false);
                var oneMat = matStair.FirstOrDefault();
                double volume = matStair != null ? UnitConverter.convertUnitsToCubicMeters(stair.GetMaterialVolume(oneMat)) : 0;

                string type = stair.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString();

                stairData.Add(new StairData
                {
                    Volume = volume,
                    Type = type,
                });
            }
            return stairData;
        }
    }
}
