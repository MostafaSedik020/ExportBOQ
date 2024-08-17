using Autodesk.Revit.DB;
using ExportBOQ.Revit.Data;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using ExportBOQ.Revit.Utils;
using OfficeOpenXml.Style;


namespace ExportBOQ.Excel
{
    public static class ExportDataToExcel
    {
        /// <summary>
        ///send the column volume to excel
        /// </summary>
        public static void SendColumnsDataToExcel(Document doc, string docTitle, string savedPath,int[] checkedNumbers)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                

                // Set some properties of the Excel document
                excelPackage.Workbook.Properties.Author = "Mostafa";
                excelPackage.Workbook.Properties.Title = "Col Report";
                excelPackage.Workbook.Properties.Subject = "Export Col Data";
                excelPackage.Workbook.Properties.Created = DateTime.Now;

                // Create the worksheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Div 3");

                //title
                //color the title
                worksheet.Cells["A1:F1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(255, 51, 63, 79);
                worksheet.Cells["A2:F2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["A2:F2"].Style.Fill.BackgroundColor.SetColor(255, 51, 63, 79);
                worksheet.Cells["A1:F1"].Style.Font.Bold = true;
                worksheet.Cells["A1:F1"].Style.Font.Color.SetColor(255, 255, 255, 255);
                worksheet.Cells["A2:F2"].Style.Font.Bold = true;
                worksheet.Cells["A2:F2"].Style.Font.Color.SetColor(255, 255, 255, 255);
                worksheet.Cells["A1:F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A2:F2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                //sizing columns
                worksheet.Column(1).Width = 7.25;
                worksheet.Column(2).Width = 47.25;
                worksheet.Column(3).Width = 7.25;
                worksheet.Column(4).Width = 10.25;
                worksheet.Column(5).Width = 12.25;
                worksheet.Column(6).Width = 16.25;

                //sizing rows
                worksheet.Row(1).Height = 23.25;
                worksheet.Row(2).Height = 23.25;
                worksheet.Row(3).Height = 19.25;
                worksheet.Row(4).Height = 75;
                worksheet.Row(5).Height = 42;
                worksheet.Row(6).Height = 32.25;
                worksheet.Row(7).Height = 42;
                worksheet.Row(8).Height = 21.75;
                worksheet.Row(9).Height = 19.25;
                worksheet.Row(10).Height = 77.25;
                worksheet.Row(11).Height = 17.25;
                worksheet.Row(12).Height = 30;
                worksheet.Row(13).Height = 108;

                #region 
                //static text discription
                worksheet.Cells["A1:A2"].Merge = true;
                worksheet.Cells["A1"].Value = "Item";

                worksheet.Cells["B1"].Value = "Description";
                worksheet.Cells["B2"].Value = "Division 03 :Concrete";

                worksheet.Cells["C1:C2"].Merge = true;
                worksheet.Cells["C1"].Value = "Unit";

                worksheet.Cells["D1:D2"].Merge = true;
                worksheet.Cells["D1"].Value = "Quantity";

                worksheet.Cells["E1"].Value = "Rate";
                worksheet.Cells["E2"].Value = "SR";

                worksheet.Cells["F1"].Value = "Amount";
                worksheet.Cells["F2"].Value = "SR";

                worksheet.Cells["B3"].Value = "GUIDELINE FOR CONTRACTOR:";
                worksheet.Cells["B3"].Style.Font.Bold = true;
                worksheet.Cells["B3"].Style.Font.UnderLine = true;

                worksheet.Cells["B4"].Value = "- All items are implemented according to the Saudi Arabia" +
                                              " code, technical specifications, drawings Attached with the" +
                                              " notes written therein and the instructions of the consultant" +
                                              " engineer.";
                worksheet.Cells["B4"].Style.WrapText = true;

                worksheet.Cells["B5"].Value = "- The technical specifications issued by the Saudi Standards" +
                                              " are referenced for the various business items according to" +
                                              " each item.";
                worksheet.Cells["B5"].Style.WrapText = true;

                worksheet.Cells["B6"].Value = "- Samples and materials should be approved before" +
                                              " supplying to the site or starting manufacturing .";
                worksheet.Cells["B6"].Style.WrapText = true;

                worksheet.Cells["B7"].Value = "- The contractor shall submit operational drawings for" +
                                              " approval by the consultant before commencing" +
                                              " manufacturing for all business items.";
                worksheet.Cells["B7"].Style.WrapText = true;

                worksheet.Cells["B8"].Value = "Section 03-30-00 Cast in Place Concrete:";
                worksheet.Cells["B8"].Style.Font.Bold = true;
                worksheet.Cells["B8"].Style.Font.UnderLine = true;

                worksheet.Cells["B9"].Value = "Plain Concrete:";
                worksheet.Cells["B9"].Style.Font.Bold = true;
                worksheet.Cells["B9"].Style.Font.UnderLine = true;

                worksheet.Cells["B10"].Value = "Ready mix plain concrete with ordinary resistant cement" +
                                              " (C 20)  with The cement content shall not be less than" +
                                              " 300kg / m3; The item includes formwork and all requirements" +
                                              " to complete the work all as indicated on drawings and" +
                                              " specifications.";
                worksheet.Cells["B10"].Style.WrapText = true;

                worksheet.Cells["A12"].Value = "2";
                worksheet.Cells["B12"].Value = "Substructure, Superstructure & On Grade Reinforced" +
                                              " Concrete:";
                worksheet.Cells["B12"].Style.Font.Bold = true;
                worksheet.Cells["B12"].Style.Font.UnderLine = true;
                worksheet.Cells["B12"].Style.WrapText = true;

                worksheet.Cells["B13"].Value = "Cast-in-place ready mix reinforced concrete with ordinary" +
                                              " Portland cement, (C 35) with The cement content shall not be" +
                                              " less than 400kg / m3; the item includes all concrete materials," +
                                              " steel reinforcement,  shuttering & formwork , movement" +
                                              " joints, joint filler and all requirements  to complete the work" +
                                              " as indicated on drawings and specifications.";
                                              
                worksheet.Cells["B13"].Style.WrapText = true;
                #endregion
                // Excel data
                int row = 11;
                int sec2num = 1;
                int sec3num = 1;
                string[] names = new string[]
                {
                "Foundation", "Columns", "Retaining Walls", "Shear Walls", "Floors",
                "Slab on grade", "Stairs", "Tanks", "Stand", "Bituminous Paint", "Polyethelene Sheet", "Membrane"
                };

                // Add headers and data
                List<string> levelsReff = FilterElements.GetTheArrangedLevels(doc);
                for (int i = 0; i < checkedNumbers.Length-3; i++)
                {
                    if (checkedNumbers[i] != 0)
                    {
                        switch (names[i])
                        {
                            case "Foundation":
                                worksheet.Cells[row, 1].Value = "1.1";
                                worksheet.Cells[row, 2].Value = "To foundation";
                                double rcFoundationVolume = FoundationData.GetFoundationData(doc, "PC").Sum(x => x.Volume);
                                worksheet.Cells[row, 3].Value = "m3";
                                worksheet.Cells[row, 4].Value = rcFoundationVolume;
                                row = row +3;

                                worksheet.Cells[row, 1].Value = $"2.{sec2num}";
                                worksheet.Cells[row, 2].Value = "To foundation";
                                double pcFoundationVolume = FoundationData.GetFoundationData(doc, "RC").Sum(x => x.Volume);
                                worksheet.Cells[row, 3].Value = "m3";
                                worksheet.Cells[row, 4].Value = pcFoundationVolume;
                                worksheet.Row(row).Height = 24.75;
                                break;

                            case "Columns":
                                worksheet.Cells[row, 1].Value = $"2.{sec2num}";
                                worksheet.Cells[row, 2].Value = "To columns";
                                double columnVolume = ColumnData.GetColumnData(doc).Sum(x => x.Volume);
                                worksheet.Cells[row, 3].Value = "m3";
                                worksheet.Cells[row, 4].Value = columnVolume;
                                worksheet.Row(row).Height = 24.75;
                                break;

                            case "Retaining Walls":
                                worksheet.Cells[row, 1].Value = $"2.{sec2num}";
                                worksheet.Cells[row, 2].Value = "To Retaining Walls";
                                double rwWallVolume = WallData.GetWallData(doc, "Retaining Wall").Sum(x => x.Volume);
                                worksheet.Cells[row, 3].Value = "m3";
                                worksheet.Cells[row, 4].Value = rwWallVolume;
                                worksheet.Row(row).Height = 24.75;
                                break;

                            case "Shear Walls":
                                worksheet.Cells[row, 1].Value = $"2.{sec2num}";
                                worksheet.Cells[row, 2].Value = "To Shear Walls and Cores";
                                double swWallVolume = WallData.GetWallData(doc, "Shear Wall").Sum(x => x.Volume);
                                worksheet.Cells[row, 3].Value = "m3";
                                worksheet.Cells[row, 4].Value = swWallVolume;
                                worksheet.Row(row).Height = 24.75;
                                break;

                            case "Floors":
                                List<FramingData> beams = FramingData.GetBeamData(doc, "Beam");
                                List<FloorData> floors = FloorData.GetFloorData(doc, "Structural Floor");
                                

                                var levels = beams.Select(l => l.Level).Union(floors.Select(l => l.Level)).Distinct();
                                List<string> sortedLevels = levels.OrderBy(level => levelsReff.IndexOf(level)).ToList();

                                foreach (var level in sortedLevels)
                                {
                                    double levelBeamVolume = beams.Where(b => b.Level == level).Sum(x => x.Volume);
                                    double levelFloorVolume = floors.Where(f => f.Level == level).Sum(x => x.Volume);
                                    double combinedVolume = levelBeamVolume + levelFloorVolume;

                                    worksheet.Cells[row, 1].Value = $"2.{sec2num}";
                                    worksheet.Cells[row, 2].Value = $" To {level.ToLower()} slabs and beams";
                                    worksheet.Cells[row, 3].Value = "m3";
                                    worksheet.Cells[row, 4].Value = combinedVolume;
                                    worksheet.Row(row).Height = 24.75;
                                    row++;
                                }
                                break;

                            case "Stairs":
                                worksheet.Cells[row, 1].Value = $"2.{sec2num}";
                                worksheet.Cells[row, 2].Value = "To stairs";
                                double stairsVolume = StairData.GetStairData(doc).Sum(x => x.Volume);
                                worksheet.Cells[row, 3].Value = "m3";
                                worksheet.Cells[row, 4].Value = stairsVolume;
                                worksheet.Row(row).Height = 24.75;
                                break;

                            case "Slab on grade":
                                worksheet.Cells[row, 1].Value = $"2.{sec2num}";
                                worksheet.Cells[row, 2].Value = "To slab on grade";
                                var sogFloors = FloorData.GetFloorData(doc, "Slab on grade");
                                double sogVolume = sogFloors.Sum(x => x.Volume);
                                worksheet.Cells[row, 3].Value = "m3";
                                worksheet.Cells[row, 4].Value = sogVolume;
                                row++;
                                sec2num++;
                                
                                worksheet.Cells[row, 1].Value = $"2.{sec2num}";
                                worksheet.Cells[row, 2].Value = "To grade beams";
                                string soglevel = sogFloors.FirstOrDefault().Level;
                                List<FramingData> gradeBeams = FramingData.GetBeamData(doc, "Beam");
                                double gradeBeamsVolume = soglevel != null ? gradeBeams.Where(b => b.Level == soglevel).Sum(x => x.Volume) : 0 ;
                                worksheet.Cells[row, 3].Value = "m3";
                                worksheet.Cells[row, 4].Value = gradeBeamsVolume;

                                worksheet.Row(row).Height = 24.75;
                                break;

                            case "Tanks":
                                worksheet.Cells[row, 1].Value = $"2.{sec2num}";
                                worksheet.Cells[row, 2].Value = "To Tanks Walls";
                                double tankWallVolume = WallData.GetWallData(doc, "Tank").Sum(x => x.Volume);
                                worksheet.Cells[row, 3].Value = "m3";
                                worksheet.Cells[row, 4].Value = tankWallVolume;
                                worksheet.Row(row).Height = 24.75;
                                break;

                             default:
                                break;

                        }
                        row++;
                        sec2num++;
                    }  
                }

                //Miscellaneous part
                worksheet.Cells[row, 1].Value = "3";
                worksheet.Cells[row, 2].Value = "Miscellaneous:";
                worksheet.Cells[row, 2].Style.Font.Bold = true;
                worksheet.Cells[row, 2].Style.Font.UnderLine = true;
                row++;

                for (int i = 8; i < checkedNumbers.Length; i++)
                {
                    if (checkedNumbers[i] != 0)
                    {
                        switch (names[i])
                        {
                            case "Stand":
                                worksheet.Cells[row, 1].Value = $"3.{sec3num}";
                                worksheet.Cells[row, 2].Value = "Upstands and Downstands";
                                double standVolume1 = FramingData.GetBeamData(doc, "Stand").Sum(x => x.Volume);
                                double standVolume2 = WallData.GetWallData(doc, "Stand").Sum(x => x.Volume);
                                worksheet.Cells[row, 3].Value = "m3";
                                worksheet.Cells[row, 4].Value = standVolume1 + standVolume2;
                                worksheet.Row(row).Height = 24.75;
                                break;
                            case "Bituminous Paint":
                                worksheet.Cells[row, 1].Value = $"3.{sec3num}";
                                worksheet.Cells[row, 2].Value = "Bituminous Paint";
                                worksheet.Cells[row, 3].Value = "m2";
                                double bitPaint = MiscData.GetBitPaint(doc);
                                worksheet.Cells[row, 4].Value = bitPaint;
                                worksheet.Row(row).Height = 24.75;
                                break;
                                
                            case "Polyethelene Sheet":
                                worksheet.Cells[row, 1].Value = $"3.{sec3num}";
                                worksheet.Cells[row, 2].Value = "Polyethelene Sheet";
                                worksheet.Cells[row, 3].Value = "m2";
                                double polyarea = MiscData.GetPolySheetArea(doc);
                                worksheet.Cells[row, 4].Value = polyarea;
                                worksheet.Row(row).Height = 24.75;
                                break;

                            case "Membrane":
                                worksheet.Cells[row, 1].Value = $"3.{sec3num}";
                                worksheet.Cells[row, 2].Value = "WaterProofing Membrane";
                                worksheet.Cells[row, 3].Value = "m2";
                                double memArea = MiscData.GetMembraneArea(doc);
                                worksheet.Cells[row, 4].Value = memArea;
                                worksheet.Row(row).Height = 24.75;
                                worksheet.Row(row).Height = 24.75;

                                break;
                        }
                        sec3num++;
                        row++;

                    }
                }

                    // Select the entire worksheet
                    ExcelRange allCells = worksheet.Cells[worksheet.Dimension.Address];

                // Set  vertical alignment to center
                allCells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                // Save the file with the naming convention
                string filename = $"{docTitle}.xlsx";
                string fullSavePath = Path.Combine(savedPath, filename);
                FileInfo fileInfo = new FileInfo(fullSavePath);
                excelPackage.SaveAs(fileInfo);

                // Open the file automatically
                Process.Start(fileInfo.FullName);
                
            }
        }
    }
}
