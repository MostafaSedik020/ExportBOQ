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
    public class ColumnData : ElementData
    {
        /// <summary>
        ///Get all the data from structural Columns in Revit
        /// </summary>
        public static List<ColumnData> GetColumnData(Document doc)
        {
            //filter columns element 
            var columns = FilterElements.GetColumnFilter(doc);

            //Initialze list of columns for the data

            var columnData = new List<ColumnData>();

            //get the column volume and convert into cubic meter
            foreach (var column in columns)
            {
                var volumeParam = column.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                double volume = volumeParam != null ? volumeParam.AsDouble() : 0;
                double convertedVolume = UnitConverter.convertUnitsToCubicMeters(volume);

                string type = column.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString();

                columnData.Add(new ColumnData
                {
                    Volume = convertedVolume,
                    Type = type,
                });
            }
            return columnData;
        }
    }
}
