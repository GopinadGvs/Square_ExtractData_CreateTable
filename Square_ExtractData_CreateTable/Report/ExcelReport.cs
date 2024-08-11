﻿using System;
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
                //BackColor = new HighlightColor() { B = 144, G = 238, R = 144 },
                BackColor = new HighlightColor() { B = 0, G = 0, R = 255 },
                WordWrap = true
            };

            List<int> rowNumbersWithOutRoad = new List<int>();

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

            int rowNumberWithRoadStart = 0;

            foreach (var item in combinedPlots)
            {
                if (!item.IsRoadAvailable)
                {
                    rowNumbersWithOutRoad.Add(rowNumberWithRoadStart);
                }

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

                rowNumberWithRoadStart++;
            }

            dt1.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,
               $"{combinedPlots.Select(x => x._PlotArea).ToArray().Sum():0.00}" ,
               $"{combinedPlots.Select(x => x._MortgageArea).ToArray().Sum():0.00}" ,
               $"{combinedPlots.Select(x => x._AmenityArea).ToArray().Sum():0.00}" });

            //Write areas in table
            //dt1.Rows.Add(new object[] {$"" ,
            //   $"" ,
            //   $"" ,
            //   $"" ,
            //   $"" ,
            //   $"" ,
            //   $"Total Site Area = {String.Format("{0:0.00}", RoundValue(SiteInfo.TotalSiteArea))}" });


            int startRow = 6;

            MyDataTable dataTable1 = new MyDataTable(dt1)
            {
                SheetName = "Meters",
                PrintHeader = false,
                StartRow = startRow
            };

            //mark rows in red color if there is no Road in any of the directions
            foreach (int rowNumber in rowNumbersWithOutRoad)
            {
                dataTable1.Rows[rowNumber].UpdateSettings = true;
                dataTable1.Rows[rowNumber].Settings = highlighter;
            }

            repo.MyDataTables.Add(dataTable1);

            var mergeSettings = new StyleSettings()
            {
                //BackColor = new HighlightColor() { B = 144, G = 238, R = 144 },
                WordWrap = false,
                HorizontalAlignment = ExcelHorizontalAlignment.Left,
                VerticalAlignment = ExcelVerticalAlignment.Center,
                FontSize = 13,
            };

            int mergeStartRow = startRow + combinedPlots.Count + 3;
            int mergeStartColumn = 3;
            int mergeEndRow = mergeStartRow;
            int numberofColumsToMerge = 5;
            int mergeEndColumn = mergeStartColumn + numberofColumsToMerge;
            int padLength = 33;

            dataTable1.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, "Total Site Area ".PadRight(padLength,'-') + "= " + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.TotalSiteArea))}", mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 2);

            dataTable1.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, "Total Plots Area ".PadRight(padLength,'-') + "= " + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.PlotsArea))}", mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable1.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Amenities Area ".PadRight(padLength,'-') + "= " + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.AmenitiesArea))}", mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable1.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Open Space Area ".PadRight(padLength,'-') + "= " + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.OpenSpaceArea))}", mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable1.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Utility Area ".PadRight(padLength) + "= " + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.UtilityArea))}", mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable1.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Internal Roads Area ".PadRight(padLength) + "= " + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.InternalRoadsArea))}", mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable1.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Left Over Owner Land Area ".PadRight(padLength) + "= " + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.LeftOverOwnerLandArea))}", mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable1.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Road Widening Area ".PadRight(padLength) + "= " + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.RoadWideningArea))}", mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable1.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Splay Area ".PadRight(padLength) + "= " + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.SplayArea))}", mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable1.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Green Buffer Zone Area ".PadRight(padLength) + "= " + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.GreenArea))}", mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 2);

            dataTable1.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Difference Area ".PadRight(padLength) + "= " + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.differenceArea))}", mergeSettings));

            repo.GenerateReport();
        }

        public static void RowIncrement(ref int startRowNumber, ref int endRowNumber, int increment)
        {
            startRowNumber = startRowNumber + increment;
            endRowNumber = endRowNumber + increment;
        }

        public static double RoundValue(double value)
        {
            return Math.Round(value, 2);
        }
    }
}
