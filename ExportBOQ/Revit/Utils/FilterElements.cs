using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;

namespace ExportBOQ.Revit.Utils
{
    public static class FilterElements
    {
        /// <summary>
        /// Returns all structural framing elements in the given Revit document.
        /// </summary>
        /// <param name="doc">The Revit document to filter.</param>
        /// <returns>List of elements representing structural framing elements.</returns>
        public static IList<Element> GetBeamFilter(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            var beams = collector.OfCategory(BuiltInCategory.OST_StructuralFraming)
                                 .WhereElementIsNotElementType()
                                 .ToElements();

            return beams;

        }
        /// <summary>
        /// Returns all structural columns elements in the given Revit document.
        /// </summary>
        /// <param name="doc">The Revit document to filter.</param>
        /// <returns>List of elements representing structural columns elements.</returns>
        public static IList<Element> GetColumnFilter(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            var columns = collector.OfCategory(BuiltInCategory.OST_StructuralColumns)
                                   .WhereElementIsNotElementType()
                                   .ToElements();

            return columns;

        }
        /// <summary>
        /// Returns all structural columns elements in the given Revit document.
        /// </summary>
        /// <param name="doc">The Revit document to filter.</param>
        /// <returns>List of elements representing structural columns elements.</returns>
        public static IList<Element> GetFloorFilter(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            var floors = collector.OfCategory(BuiltInCategory.OST_Floors)
                                  .WhereElementIsNotElementType()
                                //  .OfClass(typeof(FamilyInstance))
                                  .ToElements();

            return floors;

        }
        /// <summary>
        /// Returns all structural columns elements in the given Revit document.
        /// </summary>
        /// <param name="doc">The Revit document to filter.</param>
        /// <returns>List of elements representing structural columns elements.</returns>
        public static IList<Element> GetWallFilter(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            var walls = collector.OfCategory(BuiltInCategory.OST_Walls)
                                  .WhereElementIsNotElementType()
                                  //  .OfClass(typeof(FamilyInstance))
                                  .ToElements();

            return walls;

        }
        /// <summary>
        /// Returns all Staris elements in the given Revit document.
        /// </summary>
        /// <param name="doc">The Revit document to filter.</param>
        /// <returns>List of elements representing Stairs class.</returns>
        public static IList<Stairs> GetStairFilter(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            var stairs = collector.OfCategory(BuiltInCategory.OST_Stairs)
                          .WhereElementIsNotElementType()
                          .OfClass(typeof(Stairs))
                          .Cast<Stairs>()
                          .ToList();

            return stairs;

        }
        /// <summary>
        /// Returns all Foundation elements in the given Revit document.
        /// </summary>
        /// <param name="doc">The Revit document to filter.</param>
        /// <returns>List of elements representing Foundation element.</returns>
        public static IList<Element> GetFoundationFilter(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            var foundations = collector.OfCategory(BuiltInCategory.OST_StructuralFoundation)
                                  .WhereElementIsNotElementType()
                                  //  .OfClass(typeof(FamilyInstance))
                                  .ToElements();

            return foundations;

        }
        /// <summary>
        /// Returns the RC material in the given Revit document.
        /// </summary>
        /// <param name="doc">The Revit document to filter.</param>
        /// <returns>Material RC.</returns>
        public static Material GetRcMat(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            var material = collector.OfCategory(BuiltInCategory.OST_Materials)
                                  .Where(f => f.Name.Contains("_R_"))
                                  .Cast<Material>()
                                  .FirstOrDefault();
                                  
            return material;
                                  
        }
        /// <summary>
        /// Returns the PC material in the given Revit document.
        /// </summary>
        /// <param name="doc">The Revit document to filter.</param>
        /// <returns>Material PC.</returns>
        public static Material GetPcMat(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            var material = collector.OfCategory(BuiltInCategory.OST_Materials)
                                  .Where(f => f.Name.Contains("_P_"))
                                  .Cast<Material>()
                                  .FirstOrDefault();

            return material;

        }
        /// <summary>
        /// Arrange the levels as per its elevation.
        /// </summary>
        /// <param name="doc">The Revit document to filter.</param>
        /// <returns>List of levels as string.</returns>
        public static List<string> GetTheArrangedLevels(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            var levels = collector.OfClass(typeof(Level))
                                 .Cast<Level>()
                                 .OrderBy(level => level.Elevation)
                                 .Select(level => level.Name)
                                 .ToList();
            return levels;
        }

    }
    }
