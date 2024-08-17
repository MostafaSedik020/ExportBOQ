using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ExportBOQ.Revit.Utils;

namespace ExportBOQ.Revit.Data
{
    public class FramingData :ElementData
    {
        /// <summary>
        ///Get all the data from beams data in Revit
        /// </summary>
        public static List<FramingData> GetBeamData(Document doc,string usageParam)
        {
            //filter beams element usage only from other Usages
            var beams = FilterElements.GetBeamFilter(doc)
                                         .Where(b => b.LookupParameter("USAGE").AsValueString() == usageParam)
                                         .ToList();


            //Initialze list of beams for the data

            var beamData = new List<FramingData>();

            //get the beams volume and convert into cubic meter
            foreach (var beam in beams)
            {
                var volumeParam = beam.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                double volume = volumeParam != null ? volumeParam.AsDouble() : 0;
                double convertedVolume = UnitConverter.convertUnitsToCubicMeters(volume);

                string type = beam.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString();

                string usage = beam.LookupParameter("USAGE").AsValueString();

                string level = beam.LookupParameter("Reference Level").AsValueString();
                
                beamData.Add(new FramingData
                {
                    Volume = convertedVolume,
                    Type = type,
                    Usage = usage,
                    Level = level

                });

                
            }
            return beamData;
        }
        
    }
}
