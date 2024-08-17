using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportBOQ.Revit.Utils
{
    public class ElementData
    {
        public double Volume { get; set; }

        public string Type { get; set; }

        public string Usage { get; set; }

        public string Material { get; set; }

        public string Level { get; set; }

        public double Area { get; set; }
    }
}
