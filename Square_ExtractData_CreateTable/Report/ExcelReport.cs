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
        public static void WritetoExcel(string prefix, string path, List<Plot> combinedPlots, bool printMortageInReport)
        {
            double TotalPlotsArea = combinedPlots.Select(x => x._PlotArea).ToArray().Sum();
            double TotalMortgageArea = combinedPlots.Select(x => x._MortgageArea).ToArray().Sum();

            AreaConstants areaConstants = new AreaConstants(SiteInfo.TotalSiteArea);
            areaConstants.TotalPlotArea = TotalPlotsArea;

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
            List<int> rowNumbersWithTotalSVnoAreaMismatch = new List<int>();

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
                //if no road is available or any of the boundary info is not available then add row number to list
                if (!item.IsRoadAvailable || item._NorthInfo == "-" || item._SouthInfo == "-" || item._EastInfo == "-" || item._WestInfo == "-")
                {
                    rowNumbersWithOutRoad.Add(rowNumberWithRoadStart);
                }

                List<string> combinedText = new List<string>();
                double totalAreainSVNo = 0;
                foreach (SurveyNo svno in item._ParentSurveyNos)
                {
                    //condition to eliminate 0 areas in some survey no's ex: plot no.65
                    if (item.AreaInSurveyNo[svno] > Constants.minArea)
                        combinedText.Add($"{svno.DocumentNo + "-" + svno._SurveyNo + "-" + String.Format("{0:0.00}", item.AreaInSurveyNo[svno]) + "-" + svno.LandLordName }");

                    totalAreainSVNo += item.AreaInSurveyNo[svno];
                }

                //adding row numbers with total area in surveyno mismatches with plot area to color the cell 
                double areaDifference = Math.Abs(totalAreainSVNo - item._Area);
                if (areaDifference >= Constants.areaTolerance)
                {
                    rowNumbersWithTotalSVnoAreaMismatch.Add(rowNumberWithRoadStart);
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



            int startRow = 6;

            MyDataTable dataTable1 = new MyDataTable(dt1)
            {
                SheetName = "Meters",
                PrintHeader = false,
                StartRow = startRow
            };

            //mark cells in red color if there is no Road in any of the directions
            foreach (int rowNumber in rowNumbersWithOutRoad)
            {
                //highlight entire row
                //dataTable1.Rows[rowNumber].UpdateSettings = true;
                //dataTable1.Rows[rowNumber].Settings = highlighter;

                //cell values hard coded for Boundary cells (East,South,West,North)
                List<int> cellValues = new List<int>() { 9, 10, 11, 12 };
                foreach (int cellValue in cellValues)
                {
                    dataTable1.Rows[rowNumber].Cells[cellValue].UpdateSettings = true;
                    dataTable1.Rows[rowNumber].Cells[cellValue].Settings = highlighter;
                }
            }

            //mark cells in red color with mismatch total survey no areas with plot area
            foreach (int rowNumberWithTotalSVnoAreaMismatch in rowNumbersWithTotalSVnoAreaMismatch)
            {
                dataTable1.Rows[rowNumberWithTotalSVnoAreaMismatch].Cells[8].UpdateSettings = true;
                dataTable1.Rows[rowNumberWithTotalSVnoAreaMismatch].Cells[8].Settings = highlighter;
            }

            repo.MyDataTables.Add(dataTable1);

            //create another data table
            System.Data.DataTable dt2 = new System.Data.DataTable();
            dt2.Columns.Add("Plot Number");
            dt2.Columns.Add("East");
            dt2.Columns.Add("South");
            dt2.Columns.Add("West");
            dt2.Columns.Add("North");
            dt2.Columns.Add("Plot Area");
            dt2.Columns.Add("Mortgage Plots");
            dt2.Columns.Add("Amenity Plots");
            dt2.Columns.Add("Doc.No/R.S.No./Area/Name");
            dt2.Columns.Add("EastI");
            dt2.Columns.Add("SouthI");
            dt2.Columns.Add("WestI");
            dt2.Columns.Add("NorthI");

            dt2.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,$"" ,
               $"Total Layout Area".PadRight(33) , "-" , $"{String.Format("{0:0.00}", RoundValue(SiteInfo.TotalSiteArea))}" });

            dt2.Rows.Add(new object[] { $"" });

            if (printMortageInReport)
            {
                dt2.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,$"" ,
               $"Total Plots Area".PadRight(33) , "-" , $"{String.Format("{0:0.00}", RoundValue(TotalPlotsArea)) + " (" + $"{String.Format("{0:0.00}", RoundValue(TotalPlotsArea/SiteInfo.TotalSiteArea * 100))}" + "%)" }" });

                dt2.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,$"" ,
               $"Total Mortgages Area".PadRight(33) , "-" , $"{String.Format("{0:0.00}", RoundValue(TotalMortgageArea)) + " (" + $"{String.Format("{0:0.00}", RoundValue(TotalMortgageArea/SiteInfo.TotalSiteArea * 100))}" + "%)" }" });
            }

            else
            {
                dt2.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,$"" ,
               $"Total Plots Area Including Mortgages".PadRight(33) , "-" , $"{String.Format("{0:0.00}", RoundValue(SiteInfo.PlotsArea)) + " (" + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.PlotsArea/SiteInfo.TotalSiteArea * 100))}" + "%)" }" });
            }

            dt2.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,$"" ,
               $"Total Amenities Area ".PadRight(33) , "-" , $"{String.Format("{0:0.00}", RoundValue(SiteInfo.AmenitiesArea)) + " (" + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.AmenitiesArea/SiteInfo.TotalSiteArea*100))}" + "%)" }" });

            dt2.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,$"" ,
               $"Total openspace Area ".PadRight(33) , "-" , $"{String.Format("{0:0.00}", RoundValue(SiteInfo.OpenSpaceArea)) + " (" + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.OpenSpaceArea/SiteInfo.TotalSiteArea*100))}" + "%)" }" });

            dt2.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,$"" ,
               $"Total utility Area ".PadRight(33) , "-" , $"{String.Format("{0:0.00}", RoundValue(SiteInfo.UtilityArea)) + " (" + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.UtilityArea/SiteInfo.TotalSiteArea*100))}" + "%)" }" });

            dt2.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,$"" ,
               $"Total roads Area ".PadRight(33) , "-" , $"{String.Format("{0:0.00}", RoundValue(SiteInfo.InternalRoadsArea)) + " (" + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.InternalRoadsArea/SiteInfo.TotalSiteArea*100))}" + "%)" }" });

            dt2.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,$"" ,
               $"Total leftoverland Area ".PadRight(33) , "-" , $"{String.Format("{0:0.00}", RoundValue(SiteInfo.LeftOverOwnerLandArea)) + " (" + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.LeftOverOwnerLandArea/SiteInfo.TotalSiteArea*100))}" + "%)" }" });


            dt2.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,$"" ,
               $"Total Road widening Area ".PadRight(33) , "-" , $"{String.Format("{0:0.00}", RoundValue(SiteInfo.RoadWideningArea)) + " (" + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.RoadWideningArea/SiteInfo.TotalSiteArea*100))}" + "%)" }" });

            dt2.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,$"" ,
               $"Total Splay Area ".PadRight(33) , "-" , $"{String.Format("{0:0.00}", RoundValue(SiteInfo.SplayArea)) + " (" + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.SplayArea/SiteInfo.TotalSiteArea*100))}" + "%)" }" });

            dt2.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,$"" ,
               $"Total Amenities Area ".PadRight(33) , "-" , $"{String.Format("{0:0.00}", RoundValue(SiteInfo.GreenArea)) + " (" + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.GreenArea/SiteInfo.TotalSiteArea*100))}" + "%)" }" });

            dt2.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,$"" ,
               $"Total Verified Area".PadRight(33) , "-" , $"{String.Format("{0:0.00}", RoundValue(SiteInfo.VerifiedArea)) + " (" + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.VerifiedArea/SiteInfo.TotalSiteArea*100))}" + "%)" }" });

            dt2.Rows.Add(new object[] { $"" });

            dt2.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,$"" ,
               $"Total Difference Area".PadRight(33) , "-" , $"{String.Format("{0:0.00}", RoundValue(SiteInfo.differenceArea)) + " (" + $"{String.Format("{0:0.00}", RoundValue(SiteInfo.differenceArea/SiteInfo.TotalSiteArea*100))}" + "%)" }" });

            //logic to merge columns

            int startRowfordt2 = startRow + combinedPlots.Count + 3;

            MyDataTable dataTable2 = new MyDataTable(dt2)
            {
                SheetName = "Meters",
                PrintHeader = false,
                StartRow = startRowfordt2,
                AddBorder = false
            };

            //datatable2 font change
            var highlighter2 = new StyleSettings()
            {
                //BackColor = new HighlightColor() { B = 144, G = 238, R = 144 },
                WordWrap = false,
                HorizontalAlignment = ExcelHorizontalAlignment.Center,
                VerticalAlignment = ExcelVerticalAlignment.Center,
                FontSize = 13,
            };

            var highlighter3 = new StyleSettings()
            {
                //BackColor = new HighlightColor() { B = 144, G = 238, R = 144 },
                WordWrap = false,
                HorizontalAlignment = ExcelHorizontalAlignment.Left,
                VerticalAlignment = ExcelVerticalAlignment.Center,
                FontSize = 13,
            };

            var ColorhighlighterForArea = new StyleSettings()
            {
                BackColor = new HighlightColor() { B = 0, G = 0, R = 255 },
                WordWrap = false,
                HorizontalAlignment = ExcelHorizontalAlignment.Left,
                VerticalAlignment = ExcelVerticalAlignment.Center,
                FontSize = 13,
            };

            //adjust text height of datatable2
            foreach (var row in dataTable2.Rows)
            {
                dataTable2.Rows[row.RowNumber - 1].UpdateSettings = true;
                dataTable2.Rows[row.RowNumber - 1].Settings = highlighter2;

                //make total area column in left
                dataTable2.Rows[row.RowNumber - 1].Cells[8].UpdateSettings = true;
                dataTable2.Rows[row.RowNumber - 1].Cells[8].Settings = highlighter3;
            }

            //logic added to highlight amenity, open space & utility areas as per predefined rule
            //14.12.2024 validation added for mortgage area

            int rowColorValidation = 3;
            int rowColorValidationInc = 0;

            if (printMortageInReport)
            {
                rowColorValidationInc = 1;

                if (!areaConstants.ValidateMortgage(TotalMortgageArea))
                {
                    dataTable2.Rows[rowColorValidation].Cells[8].UpdateSettings = true;
                    dataTable2.Rows[rowColorValidation].Cells[8].Settings = ColorhighlighterForArea;
                }
            }

            if (!areaConstants.ValidateAmenity(SiteInfo.AmenitiesArea))
            {
                dataTable2.Rows[rowColorValidation + rowColorValidationInc].Cells[8].UpdateSettings = true;
                dataTable2.Rows[rowColorValidation + rowColorValidationInc].Cells[8].Settings = ColorhighlighterForArea;
            }
            if (!areaConstants.ValidateOpenSpace(SiteInfo.OpenSpaceArea))
            {
                dataTable2.Rows[rowColorValidation + rowColorValidationInc + 1].Cells[8].UpdateSettings = true;
                dataTable2.Rows[rowColorValidation + rowColorValidationInc + 1].Cells[8].Settings = ColorhighlighterForArea;
            }
            if (!areaConstants.ValidateUtility(SiteInfo.UtilityArea))
            {
                dataTable2.Rows[rowColorValidation + rowColorValidationInc + 2].Cells[8].UpdateSettings = true;
                dataTable2.Rows[rowColorValidation + rowColorValidationInc + 2].Cells[8].Settings = ColorhighlighterForArea;
            }

            repo.MyDataTables.Add(dataTable2);

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
            int numberofColumsToMerge = 4;
            int mergeEndColumn = mergeStartColumn + numberofColumsToMerge;
            int padLength = 33;

            dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Layout Area".PadRight(padLength), mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 2);

            if (printMortageInReport)
            {
                dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Plots Area".PadRight(padLength), mergeSettings));

                RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

                dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Mortgages Area".PadRight(padLength), mergeSettings));
            }

            else
            {
                dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Plots Area Including Mortgages".PadRight(padLength), mergeSettings));
            }

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Amenities Area ".PadRight(padLength), mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Open Space Area ".PadRight(padLength), mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Utility Area ".PadRight(padLength), mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Internal Roads Area ".PadRight(padLength), mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Left Over Owner Land Area ".PadRight(padLength), mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Road Widening Area ".PadRight(padLength), mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Splay Area ".PadRight(padLength), mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Green Buffer Zone Area ".PadRight(padLength), mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Verified Area ".PadRight(padLength), mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 2);

            dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Total Difference Area ".PadRight(padLength), mergeSettings));

            //add missing,duplicate & other numbers to excel
            RowIncrement(ref mergeStartRow, ref mergeEndRow, 2);
            mergeEndColumn = 13;

            dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"{FindMissingNumber.missingNumbersString}", mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"Duplicate Numbers: ", mergeSettings));

            RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);

            foreach (var item in FindMissingNumber.duplicateNumbersString)
            {
                dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"{item}", mergeSettings));
                RowIncrement(ref mergeStartRow, ref mergeEndRow, 1);
            }

            dataTable2.MergeCells.Add(Tuple.Create(mergeStartRow, mergeStartColumn, mergeEndRow, mergeEndColumn, $"{FindMissingNumber.otherNumbersString}", mergeSettings));

            List<string> paths = repo.GenerateReport();

            if (!paths.Any(s => s.IndexOf("pdf", StringComparison.OrdinalIgnoreCase) >= 0))
                System.Windows.Forms.MessageBox.Show("Failed to generate Pdf..", "Square Planners Message", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

            if (!paths.Any(s => s.IndexOf("xlsx", StringComparison.OrdinalIgnoreCase) >= 0))
                System.Windows.Forms.MessageBox.Show("Failed to generate Excel..", "Square Planners Message", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
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
