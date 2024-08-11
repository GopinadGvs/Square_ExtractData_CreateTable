using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CreateExcelAndPDF;
using OfficeOpenXml.Style;

namespace Square_ExtractData_CreateTable
{
    public static class ExcelReport
    {
        public static void WritetoExcel(string prefix, string path, List<Plot> combinedPlots)
        {
            MyDataTableRepository repo = new MyDataTableRepository();
            repo.TemplatePath = Constants.ExceltemplatePath;
            repo.Prefix = prefix;
            repo.SavePath = path;
            if (!Directory.Exists(repo.SavePath))
                Directory.CreateDirectory(repo.SavePath);
            repo.MergePdf = true;
            repo.OpenExcelAfterGenerate = true;
            repo.OpenPdfAfterGenerate = true;

            var highlighter = new StyleSettings()
            {
                BackColor = new HighlightColor() { B = 144, G = 238, R = 144 },
                WordWrap = true
            };

            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("Plot Number");
            dt1.Columns.Add("East");
            dt1.Columns.Add("South");
            dt1.Columns.Add("West");
            dt1.Columns.Add("North");
            dt1.Columns.Add("Plot Area");
            dt1.Columns.Add("Mortgage Plots");
            dt1.Columns.Add("Amenity Plots");
            dt1.Columns.Add("Doc.No/R.S.No./Area/Name");
            dt1.Columns.Add("EastI");
            dt1.Columns.Add("SouthI");
            dt1.Columns.Add("WestI");
            dt1.Columns.Add("NorthI");

            foreach (var item in combinedPlots)
            {
                List<string> combinedText = new List<string>();
                foreach (SurveyNo svno in item._ParentSurveyNos)
                {
                    //condition to eliminate 0 areas in some survey no's ex: plot no.65
                    if (item.AreaInSurveyNo[svno] > Constants.minArea)
                        combinedText.Add($"{svno.DocumentNo + "-" + svno._SurveyNo + "-" + String.Format("{0:0.00}", item.AreaInSurveyNo[svno]) + "-" + svno.LandLordName }");
                }

                dt1.Rows.Add(new object[] { $"{item._PlotNo}" ,
                $"{item._SizesInEast[0].Text}" ,
                $"{item._SizesInSouth[0].Text}" ,
                $"{item._SizesInWest[0].Text}" ,
                $"{item._SizesInNorth[0].Text}" ,
                String.Format("{0:0.00}", item._PlotArea),
                String.Format("{0:0.00}", item._MortgageArea),
                String.Format("{0:0.00}", item._AmenityArea),
                $"{Convert.ToString(string.Join(", ", combinedText.ToArray()))}" ,
                $"{item._EastInfo}" ,
                $"{item._SouthInfo}" ,
                $"{item._WestInfo}" ,
                $"{item._NorthInfo}" });
            }

            dt1.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,
               $"{combinedPlots.Select(x => x._PlotArea).ToArray().Sum():0.00}" ,
               $"{combinedPlots.Select(x => x._MortgageArea).ToArray().Sum():0.00}" ,
               $"{combinedPlots.Select(x => x._AmenityArea).ToArray().Sum():0.00}" });

            //dt1.Rows.Add(new object[] {$"" ,
            //   $"" ,
            //   $"" ,
            //   $"" ,
            //   $"" ,
            //   $"Total Site Area : {SiteInfo.TotalSiteArea}" });

            int startRow = 6;

            MyDataTable dataTable1 = new MyDataTable(dt1)
            {
                SheetName = "Meters",
                PrintHeader = false,
                StartRow = startRow
            };

            repo.MyDataTables.Add(dataTable1);

            var mergeSettings = new StyleSettings()
            {
                //BackColor = new HighlightColor() { B = 144, G = 238, R = 144 },
                WordWrap = true,
                HorizontalAlignment = ExcelHorizontalAlignment.Left,
                VerticalAlignment = ExcelVerticalAlignment.Center
            };

            int mergeStartRow = startRow + combinedPlots.Count + 3;
            int mergeStartColumn = 3;
            int mergeEndRow = mergeStartRow;
            int numberofColumsToMerge = 5;
            int mergeEndColumn = mergeStartColumn + numberofColumsToMerge;

            dataTable1.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Site Area - {String.Format("{0:0.00}", Math.Round(SiteInfo.TotalSiteArea, 2))}", mergeSettings));

            dataTable1.MergeCells.Add(Tuple.Create(mergeStartRow + 2, mergeStartColumn, mergeEndRow + 2, mergeEndColumn, $"Total Plots Area - {String.Format("{0:0.00}", Math.Round(SiteInfo.PlotsArea, 2))}", mergeSettings));

            dataTable1.MergeCells.Add(Tuple.Create(mergeStartRow + 3, mergeStartColumn, mergeEndRow + 3, mergeEndColumn, $"Total Amenities Area - {String.Format("{0:0.00}", Math.Round(SiteInfo.AmenitiesArea, 2))}", mergeSettings));

            repo.GenerateReport();
        }
    }
}
