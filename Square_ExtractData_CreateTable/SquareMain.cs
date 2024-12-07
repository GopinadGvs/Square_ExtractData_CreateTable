using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Square_ExtractData_CreateTable;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Data;
using System.Runtime.InteropServices;
using System.Threading;
using Square_ExtractData_CreateTable.ViewModel;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.Colors;

[assembly: CommandClass(typeof(MyCommands))]

namespace Square_ExtractData_CreateTable
{
    public class MyCommands
    {
        private static Editor ed = null;

        //public Form1 frm;
        //private Thread cadThread1;
        //public MyWPF myWpfForm;
        //public static ViewModelObject viewModel;

        //[CommandMethod("PB")]
        public void ProgressBarManaged()

        {
            ProgressMeter pm = new ProgressMeter();

            pm.Start("Testing Progress Bar");

            pm.SetLimit(100);

            // Now our lengthy operation

            for (int i = 0; i <= 100; i++)

            {
                System.Threading.Thread.Sleep(5);

                // Increment Progress Meter...

                pm.MeterProgress();

                // This allows AutoCAD to repaint

                System.Windows.Forms.Application.DoEvents();

            }

            pm.Stop();

        }

        [CommandMethod("SC")]
        public void SurveyNumberCheck()
        {
            if (IsLicenseExpired())
                return;

            Document acDoc = Application.DocumentManager.MdiActiveDocument;

            // Set the CMDECHO system variable to 0
            Application.SetSystemVariable("CMDECHO", 0);

            Database acCurDb = acDoc.Database;
            /*Editor*/
            ed = acDoc.Editor;

            //filename
            //string logPath = acCurDb.Filename;

            // Zoom extents
            ed.Command("_.zoom", "_e");

            string surveyNoLayer = "_SurveyNo";

            ed.Command("_-layer", "t", surveyNoLayer, "ON", surveyNoLayer, "U", surveyNoLayer, "");

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                #region Create Free Space Layer

                //create free space layer with red color

                CreateLayerByName(acCurDb, acTrans, Constants.FreeSpaceLayer, Color.FromRgb(255, 0, 0));

                #endregion

                List<ObjectId> surveyNumberPolylineIds = GetPolyLines(surveyNoLayer, acTrans);
                List<(ObjectId, string)> surveyNumberTextList = GetListTextAndObjectIdFromLayer(surveyNoLayer, acTrans);

                double totalPlotArea = Math.Round(GetAreaByLayer("_Plot", acTrans), 2);
                double surveyNosPlotArea = Math.Round(GetAreaByLayer("_SurveyNo", acTrans), 2);

                //create a dictionary with item number and it's repetative count to get duplicates
                Dictionary<string, int> dictionary = new Dictionary<string, int>();
                foreach (string item2 in surveyNumberTextList.Select(x => x.Item2).ToList())
                {
                    if (dictionary.ContainsKey(item2))
                    {
                        dictionary[item2]++;
                    }
                    else
                    {
                        dictionary[item2] = 1;
                    }
                }

                //Validate for Duplicate Survey Numbers
                //get duplicate values
                List<(string, string)> duplicatesInfo = new List<(string, string)>();
                foreach (KeyValuePair<string, int> item3 in dictionary)
                {
                    if (item3.Value > 1)
                    {
                        duplicatesInfo.Add((item3.Key, $"Value {item3.Key} occurred {item3.Value} times"));
                    }
                }

                List<ObjectId> polylineIdsForMissingSurveyNos = new List<ObjectId>();
                List<(ObjectId, string)> SurveyNosForMissingpolylines = new List<(ObjectId, string)>();
                List<(ObjectId, List<string>)> polylineIdsWithMultipleSurveyNos = new List<(ObjectId, List<string>)>();

                if (surveyNumberPolylineIds.Count == surveyNumberTextList.Count)
                {
                    System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show("No Errors Found, Do you want to open Report?", "No Mistake", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);

                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {
                        CreateReport();
                    }
                }
                else
                {
                    CreateReport();
                }

                void CreateReport()
                {

                    Dictionary<ObjectId, string> surveyNumberDictionary = GetListTextDictionaryFromLayer(surveyNoLayer, acTrans, polylineIdsWithMultipleSurveyNos);

                    //Validate for missing Survey Numbers
                    if (surveyNumberPolylineIds.Count > surveyNumberTextList.Count)
                    {
                        List<ObjectId> validatedPolyLineIds = surveyNumberDictionary.Keys.ToList();

                        polylineIdsForMissingSurveyNos = surveyNumberPolylineIds.Except(validatedPolyLineIds).ToList();

                    }
                    //Validate for missing Survey Polylines
                    else if (surveyNumberPolylineIds.Count < surveyNumberTextList.Count)
                    {
                        List<(ObjectId, string)> validatedSurveyNos = surveyNumberDictionary.Select(x => (x.Key, x.Value)).ToList();

                        SurveyNosForMissingpolylines = surveyNumberTextList.Where(x => !validatedSurveyNos.Select(y => y.Item2).ToList().Contains(x.Item2)).ToList();

                        //SurveyNosForMissingpolylines = surveyNumberTextList.Except(validatedSurveyNos).ToList();
                    }

                    string txtFileNew = Path.Combine(Path.GetDirectoryName(acCurDb.Filename), Path.GetFileNameWithoutExtension(acCurDb.Filename) + ".txt");

                    using (StreamWriter sw = new StreamWriter(txtFileNew))
                    {
                        if (duplicatesInfo.Count > 0)
                            sw.WriteLine("Duplicate Survey No's :");
                        foreach (var duplicate in duplicatesInfo)
                        {
                            sw.WriteLine($"\n{duplicate.Item2}");

                            List<ObjectId> duplicates = surveyNumberTextList.Where(x => x.Item2 == duplicate.Item1).Select(x => x.Item1).ToList();

                            foreach (ObjectId duplicateId in duplicates)
                            {
                                DBObject dbObject = acTrans.GetObject(duplicateId, OpenMode.ForRead);

                                if (dbObject is MText mText)
                                    CreatePoints(new List<Point3d>() { mText.Location });
                                if (dbObject is DBText dBText)
                                    CreatePoints(new List<Point3d>() { dBText.Position });
                            }
                        }
                        if (SurveyNosForMissingpolylines.Count > 0)
                            sw.WriteLine("\nSurvey Numbers for Missing or Incorrect Polylines :");
                        foreach (var SurveyNosForMissingpolyline in SurveyNosForMissingpolylines)
                        {
                            sw.WriteLine($"\n{SurveyNosForMissingpolyline.Item2}");

                            DBObject dbObject = acTrans.GetObject(SurveyNosForMissingpolyline.Item1, OpenMode.ForRead);

                            if (dbObject is MText mText)
                                CreatePoints(new List<Point3d>() { mText.Location });
                            if (dbObject is DBText dBText)
                                CreatePoints(new List<Point3d>() { dBText.Position });

                        }
                        if (polylineIdsForMissingSurveyNos.Count > 0)
                            sw.WriteLine("\nPolyline ID's for Missing Survey No's :");
                        foreach (ObjectId polylineIdsForMissingSurveyNo in polylineIdsForMissingSurveyNos)
                        {
                            sw.WriteLine($"\n{polylineIdsForMissingSurveyNo}");

                            Polyline acPoly = acTrans.GetObject(polylineIdsForMissingSurveyNo, OpenMode.ForRead) as Polyline;

                            Point3dCollection point3DCollection = new Point3dCollection();

                            for (int i = 0; i < acPoly.NumberOfVertices; i++)
                            {
                                point3DCollection.Add(acPoly.GetPoint3dAt(i));
                            }

                            CreatePoints(point3DCollection.Cast<Point3d>().ToList());
                        }
                        //validate for multiple survey numbers in same survey Number polyline
                        if (polylineIdsWithMultipleSurveyNos.Count > 0)
                            sw.WriteLine("\nMultiple Survey Numbers in Same Polyline :");
                        foreach (var polylineIdsWithMultipleSurveyNo in polylineIdsWithMultipleSurveyNos)
                        {
                            sw.WriteLine($"\n{polylineIdsWithMultipleSurveyNo.Item1} - {string.Join(",", polylineIdsWithMultipleSurveyNo.Item2)}");

                            Polyline acPoly = acTrans.GetObject(polylineIdsWithMultipleSurveyNo.Item1, OpenMode.ForRead) as Polyline;

                            Point3dCollection point3DCollection = new Point3dCollection();

                            for (int i = 0; i < acPoly.NumberOfVertices; i++)
                            {
                                point3DCollection.Add(acPoly.GetPoint3dAt(i));
                            }

                            CreatePoints(point3DCollection.Cast<Point3d>().ToList());

                        }
                        //Validate for Total Area of Survey Polylines with Plot Layer 
                        sw.WriteLine("\nIs Total Area Matched with Plot Layer :");
                        if (Math.Abs(surveyNosPlotArea - totalPlotArea) < 0.01)
                        {
                            sw.WriteLine("\nTotal Area is Matched.");
                            sw.WriteLine($"\nTotal Area From Survey No = {surveyNosPlotArea}");
                            sw.WriteLine($"\nTotal Area From Plot Layer = {totalPlotArea}");
                        }
                        else
                        {
                            sw.WriteLine("\nTotal Area is Not Matched.");
                            sw.WriteLine($"\nTotal Area From Survey No = {surveyNosPlotArea}");
                            sw.WriteLine($"\nTotal Area From Plot Layer = {totalPlotArea}");
                        }
                    }

                    //Set Point Mode
                    Application.SetSystemVariable("PDMODE", 35);

                    //set PDSIZE to set point size
                    Application.SetSystemVariable("PDSIZE", 2.0);

                    System.Diagnostics.Process.Start(txtFileNew);
                }

                //ToDo - Today                
                acTrans.Commit();
            }

        }



        [CommandMethod("CLAYER")]
        public void CheckAndCreateLayers()
        {
            // Get the current database and transaction manager

            if (IsLicenseExpired())
                return;

            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;

            Application.SetSystemVariable("CMDECHO", 0);

            // Zoom extents
            ed.Command("_.zoom", "_e");

            List<string> layersList = GetLayerList();

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the LayerTable for read
                LayerTable layerTable = (LayerTable)acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead);

                // Upgrade the LayerTable to write
                layerTable.UpgradeOpen();

                foreach (var layerName in layersList)
                {
                    // Check if the layer exists
                    if (!layerTable.Has(layerName))
                    {
                        // Create a new layer table record
                        LayerTableRecord layerTableRecord = new LayerTableRecord
                        {
                            Name = layerName,
                            //Color = Color.FromRgb(255, 0, 0) // Red color
                        };

                        // Add the new layer to the LayerTable
                        layerTable.Add(layerTableRecord);

                        // Add the new LayerTableRecord to the transaction
                        acTrans.AddNewlyCreatedDBObject(layerTableRecord, true);

                        // Inform the user
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nLayer {layerName} has been created...");
                    }
                    else
                    {
                        // Inform the user if the layer already exists
                        //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nLayer {layerName} already exists.");
                    }
                }

                // Commit the transaction
                acTrans.Commit();
            }

            // Execute layer management commands after the transaction is committed
            foreach (var layerName in layersList)
            {
                ed.Command("_-layer", "t", layerName, "ON", layerName, "U", layerName, "");
            }

            Application.SetSystemVariable("CMDECHO", 1);
        }


        [CommandMethod("MExportHELP")]
        public void OpenDocumentPDF()
        {
            try
            {
                System.Diagnostics.Process.Start(Constants.MExportHelpPDF);
            }
            catch { }
        }


        [CommandMethod("CHECKPOLYLINES")]
        public void CheckPolylines()
        {
            if (IsLicenseExpired())
                return;

            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;

            Application.SetSystemVariable("CMDECHO", 0);

            // Zoom extents
            ed.Command("_.zoom", "_e");

            string newLayerName = "_OpenPolyLines";

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                CreateLayerByName(acCurDb, acTrans, newLayerName, Color.FromRgb(255, 0, 0));

                PromptSelectionResult acSSPrompt1 = ed.SelectAll(CreateSelectionFilterByStartType(Constants.LWPOLYLINE));

                if (acSSPrompt1.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt1.Value;

                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        if (acSSObj != null)
                        {
                            Polyline acPoly = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;

                            if (acPoly != null && !acPoly.Closed)
                            {
                                // Clone the polyline
                                Polyline clonedPolyline = acPoly.Clone() as Polyline;

                                // Set layer
                                clonedPolyline.Layer = newLayerName;

                                // Add the cloned polyline to the model space
                                BlockTable blockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                                BlockTableRecord modelSpace = acTrans.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                                modelSpace.AppendEntity(clonedPolyline);
                                acTrans.AddNewlyCreatedDBObject(clonedPolyline, true);
                            }
                        }
                    }
                }

                // Commit the transaction
                acTrans.Commit();
            }

            Application.SetSystemVariable("CMDECHO", 1);
        }

        private List<string> GetLayerList()
        {
            List<string> layersList = new List<string>()
            {
                Constants.SurveyNoLayer,
                Constants.IndivPlotLayer,
                //Constants.IndivPlotDimLayer,
                Constants.MortgageLayer,
                Constants.AmenityLayer,
                //Constants.AmenityDimLayer,
                Constants.DocNoLayer,
                Constants.LandLordLayer,
                Constants.InternalRoadLayer,
                Constants.PlotLayer,
                Constants.OpenSpaceLayer,
                Constants.UtilityLayer,
                Constants.LeftOverOwnerLandLayer,
                Constants.SideBoundaryLayer,
                Constants.MainRoadLayer,
                Constants.SplayLayer,
                Constants.RoadWideningLayer,
                Constants.GreenBufferZoneLayer,
                Constants.LandLordSubLayer,
            };

            return layersList;
        }


        private List<string> GetLayerListToValidateFreeSpace()
        {
            List<string> layersListToValidate = new List<string>()
            {
                Constants.InternalRoadLayer,
                Constants.IndivPlotLayer,
                Constants.AmenityLayer,
                Constants.PlotLayer,
                Constants.UtilityLayer,
                Constants.LeftOverOwnerLandLayer,
                Constants.SideBoundaryLayer,
                Constants.MainRoadLayer,
                Constants.SplayLayer,
                Constants.OpenSpaceLayer
            };

            return layersListToValidate;
        }

        private void CreateLayerByName(Database acCurDb, Transaction acTrans, string layerName, Color layerColor)
        {
            // Open the LayerTable for read
            LayerTable layerTable = (LayerTable)acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead);

            // Upgrade the LayerTable to write
            layerTable.UpgradeOpen();

            //string layerName = Constants.FreeSpaceLayer;

            if (!layerTable.Has(layerName))
            {
                LayerTableRecord layerTableRecord = new LayerTableRecord
                {
                    Name = layerName,
                    Color = layerColor  /*Color.FromRgb(255, 0, 0)*/ // Red color
                };

                // Add the new layer to the LayerTable
                layerTable.Add(layerTableRecord);

                // Add the new LayerTableRecord to the transaction
                acTrans.AddNewlyCreatedDBObject(layerTableRecord, true);
            }
        }


        [CommandMethod("MExport")]

        public void SIP()
        {
            if (IsLicenseExpired())
                return;

            Document acDoc = Application.DocumentManager.MdiActiveDocument;

            // Set the CMDECHO system variable to 0
            Application.SetSystemVariable("CMDECHO", 0);

            //viewModel = new ViewModelObject();
            //myWpfForm = new MyWPF(viewModel);
            //frm = new Form1(viewModel);

            //viewModel.progressMessage = "started";
            //viewModel.progressValue = 10;

            //viewModel = new ViewModel();
            //{
            //    ShowUI = ShowUI,
            //};

            //cadThread1 = new Thread(ShowUI);
            //cadThread1.Start();

            //UpdateProgressMessage(0);

            LogWriter.LogWrite(acDoc.Name);

            Database acCurDb = acDoc.Database;
            /*Editor*/
            ed = acDoc.Editor;

            // Zoom extents
            ed.Command("_.zoom", "_e");

            // Turn on and thaw layers
            //ed.Command("_-layer", "t", Constants.SurveyNoLayer, "", "ON", Constants.SurveyNoLayer, "", "t", Constants.IndivPlotLayer, "", "ON", Constants.IndivPlotLayer, "");

            //ed.WriteMessage("Displaying Layers...");

            //viewModel.progressMessage = "Displaying";
            //viewModel.progressValue = 20;
            //UpdateProgressMessage(10, "Displaying Layers...");

            ProgressMeter pm = new ProgressMeter();
            pm.Start("Export to Excel In Progress....");
            pm.SetLimit(100);

            List<string> layersList = GetLayerList();

            // Turn on, unlock and thaw layers
            foreach (var layerName in layersList)
            {
                ed.Command("_-layer", "t", layerName, "ON", layerName, "U", layerName, "");
            }

            //ed.Command("_-layer", "t", Constants.IndivPlotLayer, "ON", Constants.IndivPlotLayer, "U", Constants.IndivPlotLayer, "");
            //List<(string, ObjectId)> snoPno = new List<(string, ObjectId)>(); -> old
            //List<(string, string)> snoPnoVal = new List<(string, string)>(); -> old

            List<SurveyNo> surveyNos = new List<SurveyNo>();
            List<Mortgage> mortgages = new List<Mortgage>();
            List<Roadline> roadlines = new List<Roadline>();
            //List<OpenSpace> OpenSpaces = new List<OpenSpace>();

            List<string> mortgagePlots = new List<string>();

            Dictionary<ObjectId, string> roadlineDict = new Dictionary<ObjectId, string>();
            Dictionary<ObjectId, string> plotlineDict = new Dictionary<ObjectId, string>();

            // Get all LWPOLYLINE entities on _SurveyNo layer
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                BlockTable acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                #region Create Free Space Layer

                //create free space layer with red color

                CreateLayerByName(acCurDb, acTrans, Constants.FreeSpaceLayer, Color.FromRgb(255, 0, 0));

                //// Open the LayerTable for read
                //LayerTable layerTable = (LayerTable)acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead);

                //// Upgrade the LayerTable to write
                //layerTable.UpgradeOpen();

                //string layerName = Constants.FreeSpaceLayer;

                //if (!layerTable.Has(layerName))
                //{
                //    LayerTableRecord layerTableRecord = new LayerTableRecord
                //    {
                //        Name = layerName,
                //        Color = Color.FromRgb(255, 0, 0) // Red color
                //    };

                //    // Add the new layer to the LayerTable
                //    layerTable.Add(layerTableRecord);

                //    // Add the new LayerTableRecord to the transaction
                //    acTrans.AddNewlyCreatedDBObject(layerTableRecord, true);
                //}

                #endregion


                #region Collect Areas

                SiteInfo.TotalSiteArea = GetAreaByLayer(Constants.PlotLayer, acTrans);

                SiteInfo.PlotsArea = GetAreaByLayer(Constants.IndivPlotLayer, acTrans);
                //MortgageArea = GetAreaByLayer(Constants.MortgageLayer, acTrans);
                SiteInfo.AmenitiesArea = GetAreaByLayer(Constants.AmenityLayer, acTrans);
                SiteInfo.OpenSpaceArea = GetAreaByLayer(Constants.OpenSpaceLayer, acTrans);
                SiteInfo.UtilityArea = GetAreaByLayer(Constants.UtilityLayer, acTrans);
                SiteInfo.InternalRoadsArea = GetAreaByLayer(Constants.InternalRoadLayer, acTrans);
                SiteInfo.SplayArea = GetAreaByLayer(Constants.SplayLayer, acTrans);
                SiteInfo.LeftOverOwnerLandArea = GetAreaByLayer(Constants.LeftOverOwnerLandLayer, acTrans);
                SiteInfo.RoadWideningArea = GetAreaByLayer(Constants.RoadWideningLayer, acTrans);
                SiteInfo.GreenArea = GetAreaByLayer(Constants.GreenBufferZoneLayer, acTrans);

                SiteInfo.VerifiedArea = SiteInfo.PlotsArea + SiteInfo.AmenitiesArea + SiteInfo.OpenSpaceArea + SiteInfo.UtilityArea + SiteInfo.InternalRoadsArea + SiteInfo.SplayArea + SiteInfo.LeftOverOwnerLandArea + SiteInfo.RoadWideningArea + SiteInfo.GreenArea;

                SiteInfo.differenceArea = SiteInfo.TotalSiteArea - SiteInfo.VerifiedArea;

                #endregion

                #region Collect Plot & Amenity numbers to find missing,duplicate & other numbers

                List<string> plotnumbers = GetListTextFromLayer(Constants.IndivPlotLayer, acTrans);
                List<string> amenityNumbers = GetListTextFromLayer(Constants.AmenityLayer, acTrans);
                List<string> combinedPlotAndAmenityNumbers = plotnumbers.Concat(amenityNumbers).ToList();

                combinedPlotAndAmenityNumbers.Sort(new AlphanumericComparer());

                List<int> sortedIntegers = new List<int>();
                List<string> sortedStrings = new List<string>();

                //separate numbers and strings
                foreach (string text3 in combinedPlotAndAmenityNumbers)
                {
                    if (int.TryParse(text3, out var _))
                    {
                        sortedIntegers.Add(Convert.ToInt32(text3));
                    }
                    else
                    {
                        sortedStrings.Add(text3);
                    }
                }

                FindMissingNumber.otherNumbersString = "Other Numbers: " + string.Join(",", sortedStrings.ToArray());

                int startNum = sortedIntegers.Min();
                int endNum = sortedIntegers.Max();

                //logic to get missing numbers from range
                IEnumerable<int> missingNumbers = Enumerable.Range(startNum, endNum - startNum).Except(sortedIntegers);

                FindMissingNumber.missingNumbersString = $"Missing Numbers in Range ({startNum} - {endNum}): " + string.Join(",", missingNumbers.ToArray());

                //create a dictionary with item number and it's repetative count to get duplicates
                Dictionary<int, int> dictionary = new Dictionary<int, int>();
                foreach (int item2 in sortedIntegers)
                {
                    if (dictionary.ContainsKey(item2))
                    {
                        dictionary[item2]++;
                    }
                    else
                    {
                        dictionary[item2] = 1;
                    }
                }

                //get duplicate values
                List<string> duplicatesInfo = new List<string>();
                foreach (KeyValuePair<int, int> item3 in dictionary)
                {
                    if (item3.Value > 1)
                    {
                        duplicatesInfo.Add($"Value {item3.Key} occurred {item3.Value} times");
                    }
                }

                FindMissingNumber.duplicateNumbersString = duplicatesInfo;

                #endregion


                #region Logic to get mortgage plot numbers list
                //ed.WriteMessage("Collecting Mortgage information...");

                //UpdateProgressMessage(20, "Collecting Mortgage information...");
                //CloseProgress();

                UpdateAutoCADProgressBar(pm);


                PromptSelectionResult acSSPrompt1 = ed.SelectAll(CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, Constants.MortgageLayer));

                if (acSSPrompt1.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt1.Value;

                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        if (acSSObj != null)
                        {
                            Mortgage mortgage = new Mortgage(); //create new SurveyNo object

                            Polyline acPoly = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;

                            mortgage._Polyline = acPoly; //assign polyline

                            if (acPoly != null)
                            {
                                mortgage.Center = getCenter(mortgage._Polyline);

                                //Collect all Points
                                for (int i = 0; i < acPoly.NumberOfVertices; i++)
                                {
                                    mortgage._PolylinePoints.Add(acPoly.GetPoint3dAt(i));
                                }

                                //get plot text inside mortgage
                                List<string> plotsInsideMortgage = GetListTextFromLayer(acTrans, mortgage._PolylinePoints, Constants.TEXT, Constants.IndivPlotLayer);
                                if (plotsInsideMortgage.Count > 0)
                                {
                                    mortgage._PlotNos.AddRange(plotsInsideMortgage);
                                }
                                else
                                {
                                    mortgage._PlotNos.AddRange(GetListTextFromLayer(acTrans, mortgage._PolylinePoints, Constants.MTEXT, Constants.IndivPlotLayer));
                                }

                                #region Old code
                                //PromptSelectionResult acSSPromptText = ed.SelectCrossingPolygon(mortgage._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer(Constants.TEXT, Constants.IndivPlotLayer));

                                //if (acSSPromptText.Status == PromptStatus.OK)
                                //{
                                //    SelectionSet acSSetTexts = acSSPromptText.Value;

                                //    foreach (SelectedObject acSSetText in acSSetTexts)
                                //    {
                                //        if (acSSetText != null)
                                //        {
                                //            DBText acText = acTrans.GetObject(acSSetText.ObjectId, OpenMode.ForRead) as DBText;
                                //            string val2 = acText.TextString;
                                //            mortgage._PlotNos.Add(val2); //add plot number to mortgage plots
                                //        }
                                //    }
                                //}
                                #endregion
                            }
                            mortgages.Add(mortgage);
                        }
                    }
                }

                mortgagePlots = mortgages.SelectMany(x => x._PlotNos).Distinct().ToList();

                #endregion

                #region Logic to get roadlines & road text Dictionary
                //ed.WriteMessage("Collecting Roadlines information...");
                UpdateAutoCADProgressBar(pm);

                PromptSelectionResult roadPrompt = ed.SelectAll(CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, Constants.InternalRoadLayer));

                if (roadPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = roadPrompt.Value;

                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        if (acSSObj != null)
                        {
                            Roadline Roadline = new Roadline(); //create new roadline object

                            Polyline acPoly = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;

                            if (acPoly != null && acPoly.Closed) //take only closed polylines for roadline
                            {
                                Roadline._Polyline = acPoly; //assign polyline
                                Roadline.Center = getCenter(Roadline._Polyline);

                                //Collect all Points
                                for (int i = 0; i < acPoly.NumberOfVertices; i++)
                                {
                                    Roadline._PolylinePoints.Add(acPoly.GetPoint3dAt(i));
                                }

                                //get road text inside Roadline
                                string roadText = GetTextFromLayer(acTrans, Roadline._PolylinePoints, Constants.TEXT, Constants.InternalRoadLayer);
                                if (!string.IsNullOrEmpty(roadText))
                                {
                                    Roadline._RoadText = roadText;
                                }
                                else
                                {
                                    Roadline._RoadText = GetTextFromLayer(acTrans, Roadline._PolylinePoints, Constants.MTEXT, Constants.InternalRoadLayer);
                                }

                                roadlines.Add(Roadline);
                            }
                        }
                    }
                }

                foreach (var item in roadlines)
                {
                    roadlineDict.Add(item._Polyline.ObjectId, item._RoadText);
                }

                #endregion

                #region Commented Logic to get openSpaceDict            

                //PromptSelectionResult openSpacePrompt = ed.SelectAll(CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, Constants.OpenSpaceLayer));

                //if (openSpacePrompt.Status == PromptStatus.OK)
                //{
                //    SelectionSet acSSet = openSpacePrompt.Value;

                //    foreach (SelectedObject acSSObj in acSSet)
                //    {
                //        if (acSSObj != null)
                //        {
                //            OpenSpace openSpace = new OpenSpace(); //create new roadline object

                //            Polyline acPoly = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;

                //            if (acPoly != null && acPoly.Closed) //take only closed polylines for roadline
                //            {
                //                openSpace._Polyline = acPoly; //assign polyline
                //                openSpace.Center = getCenter(openSpace._Polyline);

                //                //Collect all Points
                //                for (int i = 0; i < acPoly.NumberOfVertices; i++)
                //                {
                //                    openSpace._PolylinePoints.Add(acPoly.GetPoint3dAt(i));
                //                }

                //                //get road text inside Roadline
                //                string roadText = GetTextFromLayer(acTrans, openSpace._PolylinePoints, Constants.TEXT, Constants.InternalRoadLayer);
                //                if (!string.IsNullOrEmpty(roadText))
                //                {
                //                    openSpace._OpenSpaceText = roadText;
                //                }
                //                else
                //                {
                //                    openSpace._OpenSpaceText = GetTextFromLayer(acTrans, openSpace._PolylinePoints, Constants.MTEXT, Constants.InternalRoadLayer);
                //                }

                //                OpenSpaces.Add(openSpace);
                //            }
                //        }
                //    }
                //}

                //foreach (var item in OpenSpaces)
                //{
                //    openSpaceDict.Add(item._Polyline.ObjectId, item._OpenSpaceText);
                //}

                #endregion

                Dictionary<ObjectId, string> openSpaceDict = GetObjectIdAndTextDictionary(Constants.OpenSpaceLayer, acTrans);
                Dictionary<ObjectId, string> utilityDict = GetObjectIdAndTextDictionary(Constants.UtilityLayer, acTrans);
                Dictionary<ObjectId, string> LeftOverLandDict = GetObjectIdAndTextDictionary(Constants.LeftOverOwnerLandLayer, acTrans);
                Dictionary<ObjectId, string> SideBoundaryDict = GetObjectIdAndTextDictionary(Constants.SideBoundaryLayer, acTrans);
                Dictionary<ObjectId, string> MainRoadDict = GetObjectIdAndTextDictionary(Constants.MainRoadLayer, acTrans);


                #region Process SurveyNo Polylines

                //ed.WriteMessage("Collecting Survey Numbers...");
                UpdateAutoCADProgressBar(pm);

                PromptSelectionResult acSSPrompt = ed.SelectAll(CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, Constants.SurveyNoLayer));

                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt.Value;

                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        if (acSSObj != null)
                        {
                            SurveyNo surveyNo = new SurveyNo(); //create new SurveyNo object

                            Polyline acPoly = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;

                            surveyNo._Polyline = acPoly; //assign polyline

                            if (acPoly != null)
                            {
                                surveyNo.Center = getCenter(surveyNo._Polyline);

                                //Collect all Points
                                for (int i = 0; i < acPoly.NumberOfVertices; i++)
                                {
                                    surveyNo._PolylinePoints.Add(acPoly.GetPoint3dAt(i));
                                }

                                #region Fill Text information inside SurveyNo Layer

                                //Get SurveyNo Text info
                                string surveyNoText = GetTextFromLayer(acTrans, surveyNo._PolylinePoints, Constants.TEXT, Constants.SurveyNoLayer);
                                if (!string.IsNullOrEmpty(surveyNoText))
                                {
                                    surveyNo._SurveyNo = surveyNoText;//assign surveyNo
                                }
                                else
                                {
                                    surveyNo._SurveyNo = GetTextFromLayer(acTrans, surveyNo._PolylinePoints, Constants.MTEXT, Constants.SurveyNoLayer);
                                }

                                //Get DocNo Text info
                                string DocNoText = GetTextFromLayer(acTrans, surveyNo._PolylinePoints, Constants.TEXT, Constants.DocNoLayer);
                                if (!string.IsNullOrEmpty(DocNoText))
                                {
                                    surveyNo.DocumentNo = DocNoText;//assign DocNo
                                }
                                else
                                {
                                    surveyNo.DocumentNo = GetTextFromLayer(acTrans, surveyNo._PolylinePoints, Constants.MTEXT, Constants.DocNoLayer);
                                }


                                //Get LandLord Text info
                                string LandLordText = GetTextFromLayer(acTrans, surveyNo._PolylinePoints, Constants.TEXT, Constants.LandLordLayer);
                                if (!string.IsNullOrEmpty(LandLordText))
                                {
                                    surveyNo.LandLordName = LandLordText;//assign LandLordName
                                }
                                else
                                {
                                    surveyNo.LandLordName = GetTextFromLayer(acTrans, surveyNo._PolylinePoints, Constants.MTEXT, Constants.LandLordLayer);
                                }

                                #endregion

                                //ed.WriteMessage("Collecting Plot Numbers...");

                                #region Collect & Fill IndivSubPlot data using Cross Polygon                                

                                PromptSelectionResult acSSPromptPoly = ed.SelectCrossingPolygon(surveyNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, Constants.IndivPlotLayer));

                                if (acSSPromptPoly.Status == PromptStatus.OK)
                                {
                                    SelectionSet acSSetPoly = acSSPromptPoly.Value;

                                    foreach (SelectedObject acSSObjPoly in acSSetPoly)
                                    {
                                        if (acSSObjPoly != null)
                                        {
                                            Plot plotNo = new Plot(); //create new PlotNo object

                                            Polyline acPoly2 = acTrans.GetObject(acSSObjPoly.ObjectId, OpenMode.ForRead) as Polyline;

                                            if (acPoly2 != null)
                                            {
                                                List<Point3d> intersectionPoints = GetIntersections(acPoly, acPoly2);
                                                List<Point3d> uniquePoints = RemoveConsecutiveDuplicates(intersectionPoints);

                                                if (uniquePoints.Count > Constants.uniquePointsIdentifier)
                                                {
                                                    // new logic added
                                                    var existingPlotNos = surveyNos.SelectMany(x => x._PlotNos).Where
                                                    (x => x._Polyline.ObjectId == acPoly2.ObjectId).ToList();

                                                    if (existingPlotNos.Count > 0)
                                                    {
                                                        plotNo = existingPlotNos[0];
                                                        plotNo._ParentSurveyNos.Add(surveyNo);

                                                        //collect all points inside surveyno & intersecting points of plot polygon
                                                        List<Point3d> InsideAndIntersectingPoints = new List<Point3d>();

                                                        foreach (Point3d point in plotNo._PolylinePoints)
                                                        {
                                                            if (IsPointInsidePolyline(point, acPoly))
                                                                InsideAndIntersectingPoints.Add(point);
                                                        }

                                                        InsideAndIntersectingPoints = InsideAndIntersectingPoints.Concat(uniquePoints).Distinct().ToList();
                                                        plotNo.pointsInSurveyNo.Add(surveyNo, InsideAndIntersectingPoints);
                                                        double areainSurveyNo = CalculateArea(InsideAndIntersectingPoints);
                                                        plotNo.AreaInSurveyNo.Add(surveyNo, areainSurveyNo);
                                                    }

                                                    else
                                                    {
                                                        //considering dimensions from layer Constants.IndivPlotDimLayer
                                                        plotNo = FillPlotObject(surveyNos, acPoly2, plotNo, surveyNo, acTrans, Constants.IndivPlotLayer/*, Constants.IndivPlotDimLayer*/);

                                                        //collect all points inside surveyno & intersecting points of plot polygon
                                                        List<Point3d> InsideAndIntersectingPoints = new List<Point3d>();

                                                        foreach (Point3d point in plotNo._PolylinePoints)
                                                        {
                                                            if (IsPointInsidePolyline(point, acPoly))
                                                                InsideAndIntersectingPoints.Add(point);
                                                        }

                                                        InsideAndIntersectingPoints = InsideAndIntersectingPoints.Concat(uniquePoints).Distinct().ToList();
                                                        plotNo.pointsInSurveyNo.Add(surveyNo, InsideAndIntersectingPoints);
                                                        double areainSurveyNo = CalculateArea(InsideAndIntersectingPoints);
                                                        plotNo.AreaInSurveyNo.Add(surveyNo, areainSurveyNo);
                                                    }

                                                    if (!string.IsNullOrEmpty(plotNo._PlotNo))
                                                        surveyNo._PlotNos.Add(plotNo); //add plotNo to SurveyNo List if it is not empty
                                                }
                                            }
                                        }
                                    }
                                }

                                #endregion

                                #region Collect & Fill IndivSubPlot data using Window Polygon

                                // Using WindowPolygon selection to find intersections
                                PromptSelectionResult acSSPromptZeroPoly = ed.SelectWindowPolygon(surveyNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, Constants.IndivPlotLayer));

                                if (acSSPromptZeroPoly.Status == PromptStatus.OK)
                                {
                                    SelectionSet acSSetZeroPoly = acSSPromptZeroPoly.Value;

                                    foreach (SelectedObject acSSObjZeroPoly in acSSetZeroPoly)
                                    {
                                        if (acSSObjZeroPoly != null)
                                        {
                                            Plot plotNo = new Plot(); //create new PlotNo object

                                            Polyline acPoly2 = acTrans.GetObject(acSSObjZeroPoly.ObjectId, OpenMode.ForRead) as Polyline;
                                            if (acPoly2 != null)
                                            {
                                                // new logic added
                                                var existingPlotNos = surveyNos.SelectMany(x => x._PlotNos).Where
                                                (x => x._Polyline.ObjectId == acPoly2.ObjectId).ToList();

                                                if (existingPlotNos.Count > 0)
                                                {
                                                    plotNo = existingPlotNos[0];
                                                    plotNo._ParentSurveyNos.Add(surveyNo);
                                                    plotNo.AreaInSurveyNo.Add(surveyNo, plotNo._Area);
                                                }

                                                else
                                                {
                                                    //considering dimensions from layer Constants.IndivPlotDimLayer
                                                    plotNo = FillPlotObject(surveyNos, acPoly2, plotNo, surveyNo, acTrans, Constants.IndivPlotLayer/*, Constants.IndivPlotDimLayer*/);
                                                    plotNo.AreaInSurveyNo.Add(surveyNo, plotNo._Area);
                                                }
                                                if (!string.IsNullOrEmpty(plotNo._PlotNo))
                                                    surveyNo._PlotNos.Add(plotNo); //add plotNo to SurveyNo List if it is not empty
                                            }
                                        }
                                    }
                                }

                                #endregion

                                //ed.WriteMessage("Collecting Amenity Information...");

                                #region Collect & Fill Amenity data using Cross Polygon                                

                                PromptSelectionResult acSSPromptPoly1 = ed.SelectCrossingPolygon(surveyNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, Constants.AmenityLayer));

                                if (acSSPromptPoly1.Status == PromptStatus.OK)
                                {
                                    SelectionSet acSSetPoly = acSSPromptPoly1.Value;

                                    foreach (SelectedObject acSSObjPoly in acSSetPoly)
                                    {
                                        if (acSSObjPoly != null)
                                        {
                                            Plot amenityPlotNo = new Plot(); //create new amenity PlotNo object

                                            Polyline acPoly2 = acTrans.GetObject(acSSObjPoly.ObjectId, OpenMode.ForRead) as Polyline;

                                            if (acPoly2 != null)
                                            {
                                                List<Point3d> intersectionPoints = GetIntersections(acPoly, acPoly2);
                                                List<Point3d> uniquePoints = RemoveConsecutiveDuplicates(intersectionPoints);

                                                if (uniquePoints.Count > Constants.uniquePointsIdentifier)
                                                {
                                                    var existingPlotNos = surveyNos.SelectMany(x => x._AmenityPlots).Where
                                                (x => x._Polyline.ObjectId == acPoly2.ObjectId).ToList();

                                                    if (existingPlotNos.Count > 0)
                                                    {
                                                        amenityPlotNo = existingPlotNos[0];
                                                        amenityPlotNo._ParentSurveyNos.Add(surveyNo);

                                                        //collect all points inside surveyno & intersecting points of plot polygon
                                                        List<Point3d> InsideAndIntersectingPoints = new List<Point3d>();

                                                        foreach (Point3d point in amenityPlotNo._PolylinePoints)
                                                        {
                                                            if (IsPointInsidePolyline(point, acPoly))
                                                                InsideAndIntersectingPoints.Add(point);
                                                        }

                                                        InsideAndIntersectingPoints = InsideAndIntersectingPoints.Concat(uniquePoints).Distinct().ToList();
                                                        amenityPlotNo.pointsInSurveyNo.Add(surveyNo, InsideAndIntersectingPoints);
                                                        double areainSurveyNo = CalculateArea(InsideAndIntersectingPoints);
                                                        amenityPlotNo.AreaInSurveyNo.Add(surveyNo, areainSurveyNo);
                                                    }

                                                    else
                                                    {
                                                        //considering dimensions from layer Constants.AmenityDimLayer
                                                        amenityPlotNo = FillPlotObject(surveyNos, acPoly2, amenityPlotNo, surveyNo, acTrans, Constants.AmenityLayer/*, Constants.AmenityDimLayer*/);
                                                        amenityPlotNo.IsAmenity = true;

                                                        //collect all points inside surveyno & intersecting points of plot polygon
                                                        List<Point3d> InsideAndIntersectingPoints = new List<Point3d>();

                                                        foreach (Point3d point in amenityPlotNo._PolylinePoints)
                                                        {
                                                            if (IsPointInsidePolyline(point, acPoly))
                                                                InsideAndIntersectingPoints.Add(point);
                                                        }

                                                        InsideAndIntersectingPoints = InsideAndIntersectingPoints.Concat(uniquePoints).Distinct().ToList();
                                                        amenityPlotNo.pointsInSurveyNo.Add(surveyNo, InsideAndIntersectingPoints);
                                                        double areainSurveyNo = CalculateArea(InsideAndIntersectingPoints);
                                                        amenityPlotNo.AreaInSurveyNo.Add(surveyNo, areainSurveyNo);

                                                    }
                                                    if (!string.IsNullOrEmpty(amenityPlotNo._PlotNo))
                                                        surveyNo._AmenityPlots.Add(amenityPlotNo); //add amenityPlot to SurveyNo List if it is not empty
                                                }
                                            }
                                        }
                                    }
                                }

                                #endregion

                                #region Collect & Fill Amenity data using Window Polygon

                                // Using WindowPolygon selection to find intersections
                                PromptSelectionResult acSSPromptZeroPoly2 = ed.SelectWindowPolygon(surveyNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, Constants.AmenityLayer));

                                if (acSSPromptZeroPoly2.Status == PromptStatus.OK)
                                {
                                    SelectionSet acSSetZeroPoly = acSSPromptZeroPoly2.Value;

                                    foreach (SelectedObject acSSObjZeroPoly in acSSetZeroPoly)
                                    {
                                        if (acSSObjZeroPoly != null)
                                        {
                                            Plot amenityPlotNo = new Plot(); //create new amenity PlotNo object

                                            Polyline acPoly2 = acTrans.GetObject(acSSObjZeroPoly.ObjectId, OpenMode.ForRead) as Polyline;
                                            if (acPoly2 != null)
                                            {
                                                var existingPlotNos = surveyNos.SelectMany(x => x._AmenityPlots).Where
                                               (x => x._Polyline.ObjectId == acPoly2.ObjectId).ToList();

                                                if (existingPlotNos.Count > 0)
                                                {
                                                    amenityPlotNo = existingPlotNos[0];
                                                    amenityPlotNo._ParentSurveyNos.Add(surveyNo);
                                                    amenityPlotNo.AreaInSurveyNo.Add(surveyNo, amenityPlotNo._Area);
                                                }

                                                else
                                                {
                                                    //considering dimensions from layer Constants.AmenityDimLayer                                                    
                                                    amenityPlotNo = FillPlotObject(surveyNos, acPoly2, amenityPlotNo, surveyNo, acTrans, Constants.AmenityLayer/*, Constants.AmenityDimLayer*/);
                                                    amenityPlotNo.IsAmenity = true;
                                                    amenityPlotNo.AreaInSurveyNo.Add(surveyNo, amenityPlotNo._Area);

                                                }
                                                if (!string.IsNullOrEmpty(amenityPlotNo._PlotNo))
                                                    surveyNo._AmenityPlots.Add(amenityPlotNo); //add amenityPlot to SurveyNo List if it is not empty
                                            }
                                        }
                                    }
                                }

                                #endregion
                            }

                            surveyNos.Add(surveyNo); //add surveyNo to list
                        }
                    }
                }

                #endregion

                //    acTrans.Commit();
                //}

                var uniquePlots = surveyNos
                                    .SelectMany(x => x._PlotNos)
                                    .GroupBy(y => y._PlotNo)
                                    .Select(g => g.First())
                                    .OrderBy(x => x, new AlphanumericPlotComparer())
                                    .ToList();

                var uniqueAmenityPlots = surveyNos
                    .SelectMany(x => x._AmenityPlots)
                    .GroupBy(y => y._PlotNo)
                    .Select(g => g.First())
                    .OrderBy(x => x, new AlphanumericPlotComparer())
                    .ToList();

                var combinedPlots = uniquePlots.Concat(uniqueAmenityPlots).Distinct().ToList();


                UpdateAutoCADProgressBar(pm);

                foreach (var item in combinedPlots)
                {
                    plotlineDict.Add(item._Polyline.ObjectId, item._PlotNo);
                }

                Dictionary<ObjectId, string> combinedDict = plotlineDict.Concat(roadlineDict).ToDictionary(x => x.Key, y => y.Value);

                UpdateAutoCADProgressBar(pm);

                //filter area based on Mortgage & Amenity
                foreach (var item in combinedPlots)
                {
                    item._PlotArea = item._Area;

                    if (mortgagePlots.Contains(item._PlotNo))
                    {
                        item.IsMortgageArea = true;
                        item._MortgageArea = item._Area;
                        item._PlotArea = 0;
                    }

                    if (item.IsAmenity)
                    {
                        item._AmenityArea = item._Area;
                        item._PlotArea = 0;
                    }

                    #region Fill East Side Info
                    //ed.WriteMessage("Filling East information...");

                    List<Point3d> EastPointsCollection = new List<Point3d>();
                    //Point3d point1 = new Point3d(item.eastPoints[0].X + 0.5, item.eastPoints[0].Y - 2, 0);
                    //Point3d point2 = new Point3d(item.eastPoints[1].X + 0.5, item.eastPoints[1].Y - 2, 0);
                    //Point3d point3 = new Point3d(item.eastPoints[0].X, item.eastPoints[0].Y - 2, 0);
                    //Point3d point4 = new Point3d(item.eastPoints[1].X, item.eastPoints[1].Y - 2, 0);

                    Point3d epoint1 = new Point3d(item.eastLineSegment[0].MidPoint.X + 0.5, item.eastLineSegment[0].MidPoint.Y + 0.5, 0);
                    Point3d epoint2 = new Point3d(item.eastLineSegment[0].MidPoint.X - 0.5, item.eastLineSegment[0].MidPoint.Y - 0.5, 0);
                    EastPointsCollection.AddRange(new List<Point3d> { epoint1, epoint2/*, point3, point4*/ });

                    List<Polyline> roadPolylinesInEast = GetPolylinesUsingCrossPolygon(EastPointsCollection, acTrans, Constants.InternalRoadLayer);
                    List<Polyline> plotPolylinesInEast = GetPolylinesUsingCrossPolygon(EastPointsCollection, acTrans, Constants.IndivPlotLayer);
                    List<Polyline> amenityPolylinesInEast = GetPolylinesUsingCrossPolygon(EastPointsCollection, acTrans, Constants.AmenityLayer);

                    List<Polyline> openSpacePolylinesInEast = GetPolylinesUsingCrossPolygon(EastPointsCollection, acTrans, Constants.OpenSpaceLayer);
                    List<Polyline> utilityPolylinesInEast = GetPolylinesUsingCrossPolygon(EastPointsCollection, acTrans, Constants.UtilityLayer);
                    List<Polyline> leftoverlandPolylinesInEast = GetPolylinesUsingCrossPolygon(EastPointsCollection, acTrans, Constants.LeftOverOwnerLandLayer);
                    List<Polyline> sideBoundaryPolylinesInEast = GetPolylinesUsingCrossPolygon(EastPointsCollection, acTrans, Constants.SideBoundaryLayer);
                    List<Polyline> mainRoadPolylinesInEast = GetPolylinesUsingCrossPolygon(EastPointsCollection, acTrans, Constants.MainRoadLayer);

                    plotPolylinesInEast.Remove(item._Polyline); //remove current plot or amenity poyline from list
                    amenityPolylinesInEast.Remove(item._Polyline); //remove current plot or amenity poyline from list
                    try
                    {
                        if (roadPolylinesInEast.Count > 0)
                        {
                            string value = combinedDict[roadPolylinesInEast[0].ObjectId];
                            item._EastInfo = FormatRoadText(value);
                            item.IsRoadAvailable = true;
                        }
                        else if (plotPolylinesInEast.Count > 0)
                        {
                            string value = combinedDict[plotPolylinesInEast[0].ObjectId];
                            item._EastInfo = FormatPlotText(value);
                        }
                        else if (amenityPolylinesInEast.Count > 0)
                        {
                            string value = combinedDict[amenityPolylinesInEast[0].ObjectId];
                            item._EastInfo = FormatAmenityText(value);
                        }
                        else if (openSpacePolylinesInEast.Count > 0)
                        {
                            string value = openSpaceDict[openSpacePolylinesInEast[0].ObjectId];
                            item._EastInfo = FormatOpenSpaceText();
                        }
                        else if (utilityPolylinesInEast.Count > 0)
                        {
                            string value = utilityDict[utilityPolylinesInEast[0].ObjectId];
                            item._EastInfo = FormatUtilityText();
                        }
                        else if (leftoverlandPolylinesInEast.Count > 0)
                        {
                            string value = LeftOverLandDict[leftoverlandPolylinesInEast[0].ObjectId];
                            item._EastInfo = FormatLeftOverOwnerLandText();
                        }
                        else if (sideBoundaryPolylinesInEast.Count > 0)
                        {
                            string value = SideBoundaryDict[sideBoundaryPolylinesInEast[0].ObjectId];
                            item._EastInfo = FormatSideBoundaryText();
                        }
                        else if (mainRoadPolylinesInEast.Count > 0)
                        {
                            string value = MainRoadDict[mainRoadPolylinesInEast[0].ObjectId];
                            item._EastInfo = FormatMainRoadText();
                        }
                    }

                    catch (System.Exception ee)
                    {
                        //key not available in dictionary
                    }

                    #endregion

                    #region Fill South Side Info
                    //ed.WriteMessage("Filling South information...");

                    List<Point3d> SouthPointsCollection = new List<Point3d>();
                    Point3d spoint1 = new Point3d(item.southLineSegment[0].MidPoint.X + 0.5, item.southLineSegment[0].MidPoint.Y + 0.5, 0);
                    Point3d spoint2 = new Point3d(item.southLineSegment[0].MidPoint.X - 0.5, item.southLineSegment[0].MidPoint.Y - 0.5, 0);
                    SouthPointsCollection.AddRange(new List<Point3d> { spoint1, spoint2 });

                    List<Polyline> roadPolylinesInSouth = GetPolylinesUsingCrossPolygon(SouthPointsCollection, acTrans, Constants.InternalRoadLayer);
                    List<Polyline> plotPolylinesInSouth = GetPolylinesUsingCrossPolygon(SouthPointsCollection, acTrans, Constants.IndivPlotLayer);
                    List<Polyline> amenityPolylinesInSouth = GetPolylinesUsingCrossPolygon(SouthPointsCollection, acTrans, Constants.AmenityLayer);

                    List<Polyline> openSpacePolylinesInSouth = GetPolylinesUsingCrossPolygon(SouthPointsCollection, acTrans, Constants.OpenSpaceLayer);
                    List<Polyline> utilityPolylinesInSouth = GetPolylinesUsingCrossPolygon(SouthPointsCollection, acTrans, Constants.UtilityLayer);
                    List<Polyline> leftoverlandPolylinesInSouth = GetPolylinesUsingCrossPolygon(SouthPointsCollection, acTrans, Constants.LeftOverOwnerLandLayer);
                    List<Polyline> sideBoundaryPolylinesInSouth = GetPolylinesUsingCrossPolygon(SouthPointsCollection, acTrans, Constants.SideBoundaryLayer);
                    List<Polyline> mainRoadPolylinesInSouth = GetPolylinesUsingCrossPolygon(SouthPointsCollection, acTrans, Constants.MainRoadLayer);

                    plotPolylinesInSouth.Remove(item._Polyline); //remove current plot or amenity poyline from list
                    amenityPolylinesInSouth.Remove(item._Polyline); //remove current plot or amenity poyline from list
                    try
                    {

                        if (roadPolylinesInSouth.Count > 0)
                        {
                            string value = combinedDict[roadPolylinesInSouth[0].ObjectId];
                            item._SouthInfo = FormatRoadText(value);
                            item.IsRoadAvailable = true;
                        }
                        else if (plotPolylinesInSouth.Count > 0)
                        {
                            string value = combinedDict[plotPolylinesInSouth[0].ObjectId];
                            item._SouthInfo = FormatPlotText(value);
                        }
                        else if (amenityPolylinesInSouth.Count > 0)
                        {
                            string value = combinedDict[amenityPolylinesInSouth[0].ObjectId];
                            item._SouthInfo = FormatAmenityText(value);
                        }
                        else if (openSpacePolylinesInSouth.Count > 0)
                        {
                            string value = openSpaceDict[openSpacePolylinesInSouth[0].ObjectId];
                            item._SouthInfo = FormatOpenSpaceText();
                        }
                        else if (utilityPolylinesInSouth.Count > 0)
                        {
                            string value = utilityDict[utilityPolylinesInSouth[0].ObjectId];
                            item._SouthInfo = FormatUtilityText();
                        }
                        else if (leftoverlandPolylinesInSouth.Count > 0)
                        {
                            string value = LeftOverLandDict[leftoverlandPolylinesInSouth[0].ObjectId];
                            item._SouthInfo = FormatLeftOverOwnerLandText();
                        }
                        else if (sideBoundaryPolylinesInSouth.Count > 0)
                        {
                            string value = SideBoundaryDict[sideBoundaryPolylinesInSouth[0].ObjectId];
                            item._SouthInfo = FormatSideBoundaryText();
                        }
                        else if (mainRoadPolylinesInSouth.Count > 0)
                        {
                            string value = MainRoadDict[mainRoadPolylinesInSouth[0].ObjectId];
                            item._SouthInfo = FormatMainRoadText();
                        }
                    }

                    catch (System.Exception ee)
                    {
                        //key not available in dictionary
                    }

                    #endregion

                    #region Fill West Side Info
                    //ed.WriteMessage("Filling West information...");

                    List<Point3d> WestPointsCollection = new List<Point3d>();
                    Point3d wpoint1 = new Point3d(item.westLineSegment[0].MidPoint.X + 0.5, item.westLineSegment[0].MidPoint.Y + 0.5, 0);
                    Point3d wpoint2 = new Point3d(item.westLineSegment[0].MidPoint.X - 0.5, item.westLineSegment[0].MidPoint.Y - 0.5, 0);
                    WestPointsCollection.AddRange(new List<Point3d> { wpoint1, wpoint2 });

                    List<Polyline> roadPolylinesInWest = GetPolylinesUsingCrossPolygon(WestPointsCollection, acTrans, Constants.InternalRoadLayer);
                    List<Polyline> plotPolylinesInWest = GetPolylinesUsingCrossPolygon(WestPointsCollection, acTrans, Constants.IndivPlotLayer);
                    List<Polyline> amenityPolylinesInWest = GetPolylinesUsingCrossPolygon(WestPointsCollection, acTrans, Constants.AmenityLayer);

                    List<Polyline> openSpacePolylinesInWest = GetPolylinesUsingCrossPolygon(WestPointsCollection, acTrans, Constants.OpenSpaceLayer);
                    List<Polyline> utilityPolylinesInWest = GetPolylinesUsingCrossPolygon(WestPointsCollection, acTrans, Constants.UtilityLayer);
                    List<Polyline> leftoverlandPolylinesInWest = GetPolylinesUsingCrossPolygon(WestPointsCollection, acTrans, Constants.LeftOverOwnerLandLayer);
                    List<Polyline> sideBoundaryPolylinesInWest = GetPolylinesUsingCrossPolygon(WestPointsCollection, acTrans, Constants.SideBoundaryLayer);
                    List<Polyline> mainRoadPolylinesInWest = GetPolylinesUsingCrossPolygon(WestPointsCollection, acTrans, Constants.MainRoadLayer);

                    plotPolylinesInWest.Remove(item._Polyline); //remove current plot or amenity poyline from list
                    amenityPolylinesInWest.Remove(item._Polyline); //remove current plot or amenity poyline from list

                    try
                    {
                        if (roadPolylinesInWest.Count > 0)
                        {
                            string value = combinedDict[roadPolylinesInWest[0].ObjectId];
                            item._WestInfo = FormatRoadText(value);
                            item.IsRoadAvailable = true;
                        }
                        else if (plotPolylinesInWest.Count > 0)
                        {
                            string value = combinedDict[plotPolylinesInWest[0].ObjectId];
                            item._WestInfo = FormatPlotText(value);
                        }
                        else if (amenityPolylinesInWest.Count > 0)
                        {
                            string value = combinedDict[amenityPolylinesInWest[0].ObjectId];
                            item._WestInfo = FormatAmenityText(value);
                        }
                        else if (openSpacePolylinesInWest.Count > 0)
                        {
                            string value = openSpaceDict[openSpacePolylinesInWest[0].ObjectId];
                            item._WestInfo = FormatOpenSpaceText();
                        }
                        else if (utilityPolylinesInWest.Count > 0)
                        {
                            string value = utilityDict[utilityPolylinesInWest[0].ObjectId];
                            item._WestInfo = FormatUtilityText();
                        }
                        else if (leftoverlandPolylinesInWest.Count > 0)
                        {
                            string value = LeftOverLandDict[leftoverlandPolylinesInWest[0].ObjectId];
                            item._WestInfo = FormatLeftOverOwnerLandText();
                        }
                        else if (sideBoundaryPolylinesInWest.Count > 0)
                        {
                            string value = SideBoundaryDict[sideBoundaryPolylinesInWest[0].ObjectId];
                            item._WestInfo = FormatSideBoundaryText();
                        }
                        else if (mainRoadPolylinesInWest.Count > 0)
                        {
                            string value = MainRoadDict[mainRoadPolylinesInWest[0].ObjectId];
                            item._WestInfo = FormatMainRoadText();
                        }
                    }

                    catch (System.Exception ee)
                    {
                        //key not available in dictionary
                    }

                    #endregion

                    #region Fill North Side Info
                    //ed.WriteMessage("Filling North information...");

                    List<Point3d> NorthPointsCollection = new List<Point3d>();
                    Point3d npoint1 = new Point3d(item.northLineSegment[0].MidPoint.X + 0.5, item.northLineSegment[0].MidPoint.Y + 0.5, 0);
                    Point3d npoint2 = new Point3d(item.northLineSegment[0].MidPoint.X - 0.5, item.northLineSegment[0].MidPoint.Y - 0.5, 0);
                    NorthPointsCollection.AddRange(new List<Point3d> { npoint1, npoint2 });

                    List<Polyline> roadPolylinesInNorth = GetPolylinesUsingCrossPolygon(NorthPointsCollection, acTrans, Constants.InternalRoadLayer);
                    List<Polyline> plotPolylinesInNorth = GetPolylinesUsingCrossPolygon(NorthPointsCollection, acTrans, Constants.IndivPlotLayer);
                    List<Polyline> amenityPolylinesInNorth = GetPolylinesUsingCrossPolygon(NorthPointsCollection, acTrans, Constants.AmenityLayer);

                    List<Polyline> openSpacePolylinesInNorth = GetPolylinesUsingCrossPolygon(NorthPointsCollection, acTrans, Constants.OpenSpaceLayer);
                    List<Polyline> utilityPolylinesInNorth = GetPolylinesUsingCrossPolygon(NorthPointsCollection, acTrans, Constants.UtilityLayer);
                    List<Polyline> leftoverlandPolylinesInNorth = GetPolylinesUsingCrossPolygon(NorthPointsCollection, acTrans, Constants.LeftOverOwnerLandLayer);
                    List<Polyline> sideBoundaryPolylinesInNorth = GetPolylinesUsingCrossPolygon(NorthPointsCollection, acTrans, Constants.SideBoundaryLayer);
                    List<Polyline> mainRoadPolylinesInNorth = GetPolylinesUsingCrossPolygon(NorthPointsCollection, acTrans, Constants.MainRoadLayer);

                    plotPolylinesInNorth.Remove(item._Polyline); //remove current plot or amenity poyline from list
                    amenityPolylinesInNorth.Remove(item._Polyline); //remove current plot or amenity poyline from list

                    try
                    {
                        if (roadPolylinesInNorth.Count > 0)
                        {
                            string value = combinedDict[roadPolylinesInNorth[0].ObjectId];
                            item._NorthInfo = FormatRoadText(value);
                            item.IsRoadAvailable = true;
                        }
                        else if (plotPolylinesInNorth.Count > 0)
                        {
                            string value = combinedDict[plotPolylinesInNorth[0].ObjectId];
                            item._NorthInfo = FormatPlotText(value);
                        }
                        else if (amenityPolylinesInNorth.Count > 0)
                        {
                            string value = combinedDict[amenityPolylinesInNorth[0].ObjectId];
                            item._NorthInfo = FormatAmenityText(value);
                        }
                        else if (openSpacePolylinesInNorth.Count > 0)
                        {
                            string value = openSpaceDict[openSpacePolylinesInNorth[0].ObjectId];
                            item._NorthInfo = FormatOpenSpaceText();
                        }
                        else if (utilityPolylinesInNorth.Count > 0)
                        {
                            string value = utilityDict[utilityPolylinesInNorth[0].ObjectId];
                            item._NorthInfo = FormatUtilityText();
                        }
                        else if (leftoverlandPolylinesInNorth.Count > 0)
                        {
                            string value = LeftOverLandDict[leftoverlandPolylinesInNorth[0].ObjectId];
                            item._NorthInfo = FormatLeftOverOwnerLandText();
                        }
                        else if (sideBoundaryPolylinesInNorth.Count > 0)
                        {
                            string value = SideBoundaryDict[sideBoundaryPolylinesInNorth[0].ObjectId];
                            item._NorthInfo = FormatSideBoundaryText();
                        }
                        else if (mainRoadPolylinesInNorth.Count > 0)
                        {
                            string value = MainRoadDict[mainRoadPolylinesInNorth[0].ObjectId];
                            item._NorthInfo = FormatMainRoadText();
                        }
                    }

                    catch (System.Exception ee)
                    {
                        //key not available in dictionary
                    }

                    #endregion

                    #region Create points in the free space to identify gaps from individualplot & landlord sub layers if total area of landlord sub's is not matching with individual plot area

                    CreatePoints(item);

                    #endregion

                    #region Create points where east,south,west,north info is "-"

                    if (item._EastInfo.Equals("-"))
                    {

                        List<Point3d> points = new List<Point3d>();
                        points.Add(item.eastLineSegment[0].StartPoint);
                        points.Add(item.eastLineSegment[0].EndPoint);

                        CreatePoints(points);
                    }

                    if (item._SouthInfo.Equals("-"))
                    {

                        List<Point3d> points = new List<Point3d>();
                        points.Add(item.southLineSegment[0].StartPoint);
                        points.Add(item.southLineSegment[0].EndPoint);

                        CreatePoints(points);
                    }

                    if (item._WestInfo.Equals("-"))
                    {

                        List<Point3d> points = new List<Point3d>();
                        points.Add(item.westLineSegment[0].StartPoint);
                        points.Add(item.westLineSegment[0].EndPoint);

                        CreatePoints(points);
                    }

                    if (item._NorthInfo.Equals("-"))
                    {

                        List<Point3d> points = new List<Point3d>();
                        points.Add(item.northLineSegment[0].StartPoint);
                        points.Add(item.northLineSegment[0].EndPoint);

                        CreatePoints(points);
                    }

                    #endregion

                }


                #region Old logics

                //// Get unique snoPno values and sort
                //var snoPnoUnique = snoPno.Distinct().OrderBy(x => x.Item1).ToList();

                //foreach (var val in snoPnoUnique)
                //{
                //    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                //    {
                //        Entity ent = acTrans.GetObject(val.Item2, OpenMode.ForRead) as Entity;
                //        if (ent is Polyline poly)
                //        {
                //            Point3dCollection fpts = new Point3dCollection();
                //            for (int i = 0; i < poly.NumberOfVertices; i++)
                //            {
                //                fpts.Add(poly.GetPoint3dAt(i));
                //            }

                //            // Perform a selection using the window polygon method with the extracted points
                //            TypedValue[] textFilter = {
                //                new TypedValue((int)DxfCode.Start, Constants.TEXT),
                //                new TypedValue((int)DxfCode.LayerName, Constants.IndivPlotLayer)
                //            };
                //            SelectionFilter textSelFilter = new SelectionFilter(textFilter);
                //            PromptSelectionResult textSelResult = ed.SelectWindowPolygon(fpts, textSelFilter);

                //            if (textSelResult.Status == PromptStatus.OK)
                //            {
                //                DBText textEntity = acTrans.GetObject(textSelResult.Value[0].ObjectId, OpenMode.ForRead) as DBText;
                //                if (textEntity != null)
                //                {
                //                    string fval2 = textEntity.TextString;
                //                    snoPnoVal.Add((fval2, val.Item1));
                //                }
                //            }
                //        }

                //        acTrans.Commit();
                //    }
                //}

                //snoPnoVal = snoPnoVal.OrderBy(x => int.Parse(x.Item1)).ToList();

                //// Additional logic for _IndivSubPlot layer
                //List<(ObjectId, string)> ispAllData = new List<(ObjectId, string)>();
                //List<string> ispItemname = new List<string>();
                //List<string> ispSno = new List<string>();

                //using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                //{
                //    TypedValue[] acTypValArIndiv = new TypedValue[]
                //    {
                //        new TypedValue((int)DxfCode.Start, Constants.LWPOLYLINE),
                //        new TypedValue((int)DxfCode.LayerName, Constants.IndivPlotLayer)
                //    };
                //    SelectionFilter acSelFtrIndiv = new SelectionFilter(acTypValArIndiv);
                //    PromptSelectionResult acSSPromptIndiv = ed.SelectAll(acSelFtrIndiv);

                //    if (acSSPromptIndiv.Status == PromptStatus.OK)
                //    {
                //        SelectionSet acSSetIndiv = acSSPromptIndiv.Value;

                //        foreach (SelectedObject acSSObjIndiv in acSSetIndiv)
                //        {
                //            if (acSSObjIndiv != null)
                //            {
                //                Polyline acPolyIndiv = acTrans.GetObject(acSSObjIndiv.ObjectId, OpenMode.ForRead) as Polyline;

                //                if (acPolyIndiv != null)
                //                {
                //                    Point3dCollection ptsIndiv = new Point3dCollection();
                //                    for (int i = 0; i < acPolyIndiv.NumberOfVertices; i++)
                //                    {
                //                        ptsIndiv.Add(acPolyIndiv.GetPoint3dAt(i));
                //                    }

                //                    // Find text entity on _IndivSubPlot layer
                //                    TypedValue[] acTypValArTextIndiv = new TypedValue[]
                //                    {
                //                        new TypedValue((int)DxfCode.Start, Constants.TEXT),
                //                        new TypedValue((int)DxfCode.LayerName, Constants.IndivPlotLayer)
                //                    };
                //                    SelectionFilter acSelFtrTextIndiv = new SelectionFilter(acTypValArTextIndiv);
                //                    PromptSelectionResult acSSPromptTextIndiv = ed.SelectCrossingPolygon(ptsIndiv, acSelFtrTextIndiv);

                //                    if (acSSPromptTextIndiv.Status == PromptStatus.OK)
                //                    {
                //                        SelectionSet acSSetTextIndiv = acSSPromptTextIndiv.Value;
                //                        DBText acTextIndiv = acTrans.GetObject(acSSetTextIndiv[0].ObjectId, OpenMode.ForRead) as DBText;
                //                        string val2Indiv = acTextIndiv.TextString;

                //                        ispItemname.Add(acPolyIndiv.ObjectId.ToString());
                //                        ispSno.Add(val2Indiv);
                //                        ispAllData.Add((acPolyIndiv.ObjectId, val2Indiv));
                //                    }
                //                }
                //            }
                //        }
                //    }

                //    acTrans.Commit();
                //}

                //ispSno = ispSno.OrderBy(x => int.Parse(x)).ToList();

                //List<(string, string)> ispSvnoData = new List<(string, string)>();

                //foreach (var val in ispSno)
                //{
                //    var pts1 = snoPnoVal.Where(x => x.Item1 == val).Select(x => x.Item2).ToList();

                //    string ispSvno;
                //    if (pts1.Count > 1)
                //    {
                //        ispSvno = string.Join("|", pts1);
                //    }
                //    else
                //    {
                //        ispSvno = pts1.FirstOrDefault();
                //    }

                //    ispSvnoData.Add((val, ispSvno));
                //}

                // Write data to CSV
                //string csvFileNew = Path.Combine(Path.GetDirectoryName(acCurDb.Filename), Path.GetFileNameWithoutExtension(acCurDb.Filename) + ".csv");

                //using (StreamWriter sw = new StreamWriter(csvFileNew))
                //{
                //    sw.WriteLine("Plot Number,Survey No");
                //    foreach (var itm in ispSvnoData)
                //    {
                //        sw.WriteLine($"{itm.Item1},{itm.Item2}");
                //    }
                //}

                #endregion

                #region Save Excel New Dictionary

                //Dictionary<string, string> plotNoVsSurveyNo = new Dictionary<string, string>();
                //foreach (var item in uniquePlots)
                //{
                //    plotNoVsSurveyNo.Add(item._PlotNo, string.Join("|", item._ParentSurveyNos.Select(x => x._SurveyNo).ToArray()));
                //}

                //string csvFileNew = Path.Combine(Path.GetDirectoryName(acCurDb.Filename), Path.GetFileNameWithoutExtension(acCurDb.Filename) + ".csv");
                //using (StreamWriter sw = new StreamWriter(csvFileNew))
                //{
                //    sw.WriteLine("Plot Number,Survey No,Center");
                //    foreach (var itm in plotNoVsSurveyNo)
                //    {
                //        sw.WriteLine($"{itm.Key},{itm.Value}");
                //    }
                //}

                #endregion


                #region Validate all Areas and highlight free space from surveyNo & Landlord_Sub layers

                PromptSelectionResult acSSPromptNew = ed.SelectAll(CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, Constants.SurveyNoMainLayer));

                if (acSSPromptNew.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPromptNew.Value;

                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        double SurveyNoArea = 0.0;
                        double landLordSubArea = 0.0;

                        Dictionary<string, ObjectId> myDict = new Dictionary<string, ObjectId>();
                        string surveyNoText = string.Empty;

                        if (acSSObj != null)
                        {
                            Polyline acPoly = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;

                            Point3dCollection intersectionPointsOtherthanBoundary = new Point3dCollection();

                            if (acPoly != null)
                            {
                                SurveyNoArea = acPoly.Area;

                                //Collect all Points
                                Point3dCollection point3DCollection = new Point3dCollection();
                                for (int i = 0; i < acPoly.NumberOfVertices; i++)
                                {
                                    point3DCollection.Add(acPoly.GetPoint3dAt(i));
                                }

                                surveyNoText = GetTextFromLayer(acTrans, point3DCollection, Constants.TEXT, Constants.SurveyNoMainLayer);
                                if (string.IsNullOrEmpty(surveyNoText))
                                    surveyNoText = GetTextFromLayer(acTrans, point3DCollection, Constants.MTEXT, Constants.SurveyNoMainLayer);

                                myDict.Add(surveyNoText, acPoly.ObjectId);

                                //Done - 22.08.2024 use Cross window & window polygon

                                PromptSelectionResult acSSPromptPoly = ed.SelectCrossingPolygon(point3DCollection, CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, Constants.SurveyNoLayer));

                                if (acSSPromptPoly.Status == PromptStatus.OK)
                                {
                                    SelectionSet acSSetPoly = acSSPromptPoly.Value;

                                    foreach (SelectedObject acSSObjPoly in acSSetPoly)
                                    {
                                        if (acSSObjPoly != null)
                                        {
                                            Polyline acPoly2 = acTrans.GetObject(acSSObjPoly.ObjectId, OpenMode.ForRead) as Polyline;

                                            //Collect all Points
                                            Point3dCollection point3DCollection2 = new Point3dCollection();
                                            for (int i = 0; i < acPoly2.NumberOfVertices; i++)
                                            {
                                                point3DCollection2.Add(acPoly2.GetPoint3dAt(i));
                                            }

                                            if (acPoly2 != null)
                                            {
                                                List<Point3d> intersectionPoints = GetIntersections(acPoly, acPoly2);
                                                List<Point3d> uniquePoints = RemoveConsecutiveDuplicates(intersectionPoints);

                                                if (uniquePoints.Count > Constants.uniquePointsIdentifier) //ToDo - test multiple types and confirm this logic
                                                {
                                                    string text = GetTextFromLayer(acTrans, point3DCollection2, Constants.TEXT, Constants.SurveyNoLayer);
                                                    if (string.IsNullOrEmpty(text))
                                                        text = GetTextFromLayer(acTrans, point3DCollection, Constants.MTEXT, Constants.SurveyNoLayer);

                                                    if (surveyNoText == text)
                                                    {
                                                        landLordSubArea += acPoly2.Area;
                                                    }

                                                    //add points to intersectionPointsOtherthanBoundary to draw red color boundary
                                                    foreach (Point3d uniquePoint in uniquePoints)
                                                    {
                                                        if (!point3DCollection.Contains(uniquePoint))
                                                            intersectionPointsOtherthanBoundary.Add(uniquePoint);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                PromptSelectionResult acSSPromptZeroPoly = ed.SelectWindowPolygon(point3DCollection, CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, Constants.SurveyNoLayer));

                                if (acSSPromptZeroPoly.Status == PromptStatus.OK)
                                {
                                    SelectionSet acSSetZeroPoly = acSSPromptZeroPoly.Value;

                                    foreach (SelectedObject acSSObjZeroPoly in acSSetZeroPoly)
                                    {
                                        if (acSSObjZeroPoly != null)
                                        {
                                            Polyline acPoly2 = acTrans.GetObject(acSSObjZeroPoly.ObjectId, OpenMode.ForRead) as Polyline;

                                            //Collect all Points
                                            Point3dCollection point3DCollection2 = new Point3dCollection();
                                            for (int i = 0; i < acPoly2.NumberOfVertices; i++)
                                            {
                                                point3DCollection2.Add(acPoly2.GetPoint3dAt(i));
                                            }

                                            if (acPoly2 != null)
                                            {
                                                string text = GetTextFromLayer(acTrans, point3DCollection2, Constants.TEXT, Constants.SurveyNoLayer);
                                                if (string.IsNullOrEmpty(text))
                                                    text = GetTextFromLayer(acTrans, point3DCollection, Constants.MTEXT, Constants.SurveyNoLayer);

                                                if (surveyNoText == text)
                                                {
                                                    landLordSubArea += acPoly2.Area;
                                                }
                                            }
                                        }
                                    }
                                }

                                //List<Polyline> polylines = GetPolylinesUsingCrossPolygon(point3DCollection.Cast<Point3d>().ToList(), acTrans, Constants.SurveyNoLayer);

                                //foreach (Polyline polyline in polylines)
                                //{
                                //    landLordSubArea += polyline.Area;
                                //}
                            }

                            double diffArea = landLordSubArea - SurveyNoArea;

                            //if (diffArea != 0 && diffArea < Constants.areaTolerance) //ToDo Update logic here
                            //{
                            //    //Area Mismatch
                            //    System.Diagnostics.Debug.Print(surveyNoText);

                            //    if (intersectionPointsOtherthanBoundary.Count > 0)
                            //    {
                            //        
                            //        double area = CalculateAreaAndCreatePolyline(intersectionPointsOtherthanBoundary.Cast<Point3d>().ToList());
                            //    }
                            //}
                            if (diffArea == 0)
                            {
                                //correct area
                                System.Diagnostics.Debug.Print(surveyNoText);
                            }

                            else
                            {
                                //Area Mismatch
                                System.Diagnostics.Debug.Print(surveyNoText);

                                if (intersectionPointsOtherthanBoundary.Count > 0)
                                {
                                    //22.08.2024, validate logic
                                    //create polyline code commented as some times we will get only single point, so creating points only
                                    //double area = CalculateAreaAndCreatePolyline(intersectionPointsOtherthanBoundary.Cast<Point3d>().ToList());

                                    //Create points at the unidentified areas
                                    CreatePoints(intersectionPointsOtherthanBoundary.Cast<Point3d>().ToList());
                                }
                            }
                        }
                    }
                }

                #endregion

                #region Validate all poyline points in all Layers to identify free space

                List<string> layersListToValidate = GetLayerListToValidateFreeSpace();

                List<string> layersToSkip = new List<string>()
                {
                    Constants.PlotLayer,
                    Constants.MainRoadLayer,
                    Constants.SideBoundaryLayer,
                    //Constants.SplayLayer
                };

                foreach (string currentValidateLayer in layersListToValidate)
                {
                    //Skip Plot, main Road, side boundary Layer for validation
                    if (layersToSkip.Contains(currentValidateLayer))
                    {
                        continue;
                    }

                    PromptSelectionResult acSSPromptOpenSpace = ed.SelectAll(CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, currentValidateLayer));

                    if (acSSPromptOpenSpace.Status == PromptStatus.OK)
                    {
                        SelectionSet acSSet = acSSPromptOpenSpace.Value;

                        foreach (SelectedObject acSSObj in acSSet)
                        {
                            if (acSSObj != null)
                            {
                                Polyline acPoly = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;

                                if (acPoly != null && acPoly.Closed)
                                {
                                    //Collect all Points
                                    //Point3dCollection point3DCollection = new Point3dCollection();
                                    for (int i = 0; i < acPoly.NumberOfVertices; i++)
                                    {
                                        //point3DCollection.Add(acPoly.GetPoint3dAt(i));

                                        Point3d mypoint = acPoly.GetPoint3dAt(i);

                                        List<Point3d> points = new List<Point3d>();
                                        Point3d point1 = new Point3d(mypoint.X + 0.5, mypoint.Y + 0.5, 0);
                                        Point3d point2 = new Point3d(mypoint.X - 0.5, mypoint.Y - 0.5, 0);
                                        points.AddRange(new List<Point3d> { point1, point2 });

                                        List<string> layersListforValidation = GetLayerListToValidateFreeSpace();

                                        //IndivPlotLayer can intersect with IndivPlotLayer also
                                        if (currentValidateLayer != Constants.IndivPlotLayer)
                                            layersListforValidation.Remove(currentValidateLayer);

                                        List<Polyline> boundaryPolylines = new List<Polyline>();

                                        foreach (string layer in layersListforValidation)
                                        {
                                            boundaryPolylines.AddRange(GetPolylinesUsingCrossPolygon(points, acTrans, layer));
                                        }

                                        if (!boundaryPolylines.Any())
                                            CreatePoints(new List<Point3d>() { mypoint });

                                        //List<Polyline> roadPolylines = GetPolylinesUsingCrossPolygon(points, acTrans, Constants.InternalRoadLayer);
                                        //List<Polyline> plotPolylines = GetPolylinesUsingCrossPolygon(points, acTrans, Constants.IndivPlotLayer);
                                        //List<Polyline> amenityPolylines = GetPolylinesUsingCrossPolygon(points, acTrans, Constants.AmenityLayer);
                                        //List<Polyline> LayoutPolylines = GetPolylinesUsingCrossPolygon(points, acTrans, Constants.PlotLayer);
                                        //List<Polyline> utilityPolylines = GetPolylinesUsingCrossPolygon(points, acTrans, Constants.UtilityLayer);
                                        //List<Polyline> leftOverLandPolylines = GetPolylinesUsingCrossPolygon(points, acTrans, Constants.LeftOverOwnerLandLayer);
                                        //List<Polyline> sideBoundaryPolylines = GetPolylinesUsingCrossPolygon(points, acTrans, Constants.SideBoundaryLayer);
                                        //List<Polyline> mainRoadPolylines = GetPolylinesUsingCrossPolygon(points, acTrans, Constants.MainRoadLayer);
                                        //List<Polyline> splayPolylines = GetPolylinesUsingCrossPolygon(points, acTrans, Constants.SplayLayer);

                                        //if (roadPolylines.Any() || plotPolylines.Any() || amenityPolylines.Any() || LayoutPolylines.Any() ||
                                        //    utilityPolylines.Any() || leftOverLandPolylines.Any() || sideBoundaryPolylines.Any() || mainRoadPolylines.Any() || splayPolylines.Any())
                                        //{
                                        //    //found some connected side boundary
                                        //}
                                        //else
                                        //{
                                        //    //no connected boundary found
                                        //    CreatePoints(new List<Point3d>() { mypoint });
                                        //}
                                    }
                                }
                            }
                        }
                    }

                    UpdateAutoCADProgressBar(pm);
                }
                #endregion


                #region Identify polylines that are outside the plot layer

                //ToDo

                #endregion


                // Write data to CSV

                UpdateAutoCADProgressBar(pm);

                //Set Point Mode
                Application.SetSystemVariable("PDMODE", 35);

                DateTime datetime = DateTime.Now;
                string uniqueId = String.Format("{0:00}{1:00}{2:0000}{3:00}{4:00}{5:00}{6:000}",
                    datetime.Day, datetime.Month, datetime.Year,
                    datetime.Hour, datetime.Minute, datetime.Second, datetime.Millisecond);

                //System.Windows.Forms.MessageBox.Show(uniqueId);

                string fullfileName = Path.Combine(Path.GetDirectoryName(acCurDb.Filename), Path.GetFileNameWithoutExtension(acCurDb.Filename) + $"{"_" + uniqueId}");

                string csvFileNew = Path.Combine(Path.GetDirectoryName(acCurDb.Filename), Path.GetFileNameWithoutExtension(acCurDb.Filename) + $"{"_" + uniqueId}") + ".csv";

                string prefix = Path.GetFileNameWithoutExtension(acCurDb.Filename) + "_";
                string folderPath = Path.GetDirectoryName(acCurDb.Filename);

                //Commented write data to csv as it is no longer needed
                //WritetoCSV(csvFileNew, combinedPlots);

                ed.WriteMessage("Generating Report...");

                ExcelReport.WritetoExcel(prefix, folderPath, combinedPlots);

                #region Test write

                //sw.WriteLine("Plot Number,East,South,West,North,Survey No,Center,EP1,EP2,SP1,SP2,WP1,WP2,NP1,NP2");

                //$"{Convert.ToString(string.Join("|", item._ParentSurveyNos.Select(x => x._SurveyNo).ToArray()))}," +

                //{(/*item._PlotArea == 0 ? "" : */item._PlotArea.ToString())}

                //+
                //$"{Convert.ToString(Math.Round(item.Center[0], 2) + "|" + Math.Round(item.Center[1]))}," +
                //$"{Convert.ToString(Math.Round(item.eastPoints[0][0], 2) + "|" + Math.Round(item.eastPoints[0][1], 2))}," +
                //$"{Convert.ToString(Math.Round(item.eastPoints[1][0], 2) + "|" + Math.Round(item.eastPoints[1][1], 2))}," +
                //$"{Convert.ToString(Math.Round(item.southPoints[0][0], 2) + "|" + Math.Round(item.southPoints[0][1], 2))}," +
                //$"{Convert.ToString(Math.Round(item.southPoints[1][0], 2) + "|" + Math.Round(item.southPoints[1][1], 2))}," +
                //$"{Convert.ToString(Math.Round(item.westPoints[0][0], 2) + "|" + Math.Round(item.westPoints[0][1], 2))}," +
                //$"{Convert.ToString(Math.Round(item.westPoints[1][0], 2) + "|" + Math.Round(item.westPoints[1][1], 2))}," +
                //$"{Convert.ToString(Math.Round(item.northPoints[0][0], 2) + "|" + Math.Round(item.northPoints[0][1], 2))}," +
                //$"{Convert.ToString(Math.Round(item.northPoints[1][0], 2) + "|" + Math.Round(item.northPoints[1][1], 2))}"

                #endregion


                //System.Diagnostics.Process.Start("notepad.exe", csvFileNew);
                //System.Diagnostics.Process.Start("Excel.exe", csvFileNew);

                // Turn off _SurveyNo layer
                //ed.Command("_-layer", "OFF", Constants.SurveyNoLayer, "");

                //CloseProgress();

                UpdateAutoCADProgressBar(pm);
                CloseAutoCADProgress(pm);

                // Set the CMDECHO system variable to 0
                Application.SetSystemVariable("CMDECHO", 1);

                ed.WriteMessage("\nProcess complete.");

                acTrans.Commit();
            }
        }

        public Dictionary<ObjectId, string> GetObjectIdAndTextDictionary(string layerName, Transaction acTrans)
        {
            Dictionary<ObjectId, string> myDict = new Dictionary<ObjectId, string>();

            PromptSelectionResult openSpacePrompt = ed.SelectAll(CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, layerName));

            if (openSpacePrompt.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = openSpacePrompt.Value;

                foreach (SelectedObject acSSObj in acSSet)
                {
                    if (acSSObj != null)
                    {
                        Polyline acPoly = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;

                        if (acPoly != null && acPoly.Closed) //take only closed polylines for roadline
                        {
                            Point3dCollection point3DCollection = new Point3dCollection();

                            //Collect all Points
                            for (int i = 0; i < acPoly.NumberOfVertices; i++)
                            {
                                point3DCollection.Add(acPoly.GetPoint3dAt(i));
                            }

                            //get road text inside Roadline
                            string roadText = GetTextFromLayer(acTrans, point3DCollection, Constants.TEXT, Constants.InternalRoadLayer);
                            if (!string.IsNullOrEmpty(roadText))
                            {
                                myDict.Add(acPoly.ObjectId, roadText);
                            }
                            else
                            {
                                myDict.Add(acPoly.ObjectId, GetTextFromLayer(acTrans, point3DCollection, Constants.MTEXT, Constants.InternalRoadLayer));
                            }
                        }
                    }
                }
            }

            return myDict;
        }

        public string FormatRoadText(string roadText)
        {
            if (string.IsNullOrEmpty(roadText))
                return "-";

            //int len = roadText.IndexOf(" ");
            //string formattedtext = roadText.Substring(0, len > 0 ? len : 10).Trim() + " Mts. Road";

            string formattedtext = ExtractNumbersFromString(roadText)[0] + " Mts. Road";
            return formattedtext;
        }

        public string FormatPlotText(string plotText)
        {
            if (string.IsNullOrEmpty(plotText))
                return "-";

            string formattedtext = "Plot No. " + plotText;
            return formattedtext;
        }

        public string FormatAmenityText(string amenityText)
        {
            if (string.IsNullOrEmpty(amenityText))
                return "-";

            string formattedtext = "Amenity " + amenityText;
            return formattedtext;
        }

        public string FormatOpenSpaceText()
        {
            return "Open Space";
        }

        public string FormatUtilityText()
        {
            return "Utility";
        }

        public string FormatLeftOverOwnerLandText()
        {
            return "Left Over Land";
        }

        public string FormatSideBoundaryText()
        {
            return "Side Boundary";
        }

        public string FormatMainRoadText()
        {
            return "Main Road";
        }

        static List<string> ExtractNumbersFromString(string input)
        {
            // Regular expression to match numbers (including decimals)
            Regex regex = new Regex(@"\d+(\.\d+)?");

            // Find all matches
            MatchCollection matches = regex.Matches(input);

            // Collect all matches into a list of strings
            List<string> numbers = new List<string>();
            foreach (Match match in matches)
            {
                numbers.Add(match.Value);
            }

            return numbers;
        }

        public bool IsLicenseExpired()
        {
            // Retrieve the expiry date from the registry
            RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\SquarePlanner");
            if (key != null)
            {
                string expiryDateString = key.GetValue("Expiry") as string;
                if (DateTime.TryParse(expiryDateString, out DateTime expiryDate))
                {
                    // Check if the current date is past the expiry date
                    if (DateTime.Now > expiryDate)
                    {
                        return true; // License has expired
                    }
                    return false; // License is still valid
                }
            }
            return true; // Default to expired if there's an issue
        }

        public void ShowUI()
        {
            //frm.TopMost = true;
            //frm.ShowDialog();
            //myWpfForm.Show();
        }

        public void UpdateProgressMessage(int value, string text = "Square Program Started...")
        {
            //frm.toolStripStatusLabel1.Text = text;
            //System.Threading.Thread.Sleep(5);
            //frm.toolStripProgressBar1.Value = value;
        }

        private void CloseProgress()
        {
            //cadThread1?.Abort();
        }

        private void UpdateAutoCADProgressBar(ProgressMeter pm)
        {
            for (int i = 0; i < 8; i++)
            {
                System.Threading.Thread.Sleep(5);

                // Increment Progress Meter...

                pm.MeterProgress();

                // This allows AutoCAD to repaint

                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void CloseAutoCADProgress(ProgressMeter pm)
        {
            pm.Stop();
        }

        private void WritetoCSV(string csvFileNew, List<Plot> combinedPlots)
        {
            using (StreamWriter sw = new StreamWriter(csvFileNew))
            {
                sw.WriteLine("Plot Number,East,South,West,North,Plot Area, Mortgage Plots, Amenity Plots,Doc.No/R.S.No./Area/Name,East,South,West,North");

                foreach (var item in combinedPlots)
                {
                    List<string> combinedText = new List<string>();
                    foreach (SurveyNo svno in item._ParentSurveyNos)
                    {
                        //condition to eliminate 0 areas in some survey no's ex: plot no.65
                        if (item.AreaInSurveyNo[svno] > Constants.minArea)
                            combinedText.Add($"{svno.DocumentNo + "-" + svno._SurveyNo + "-" + String.Format("{0:0.00}", item.AreaInSurveyNo[svno]) + "-" + svno.LandLordName }");
                    }

                    string textValue1 = $"{item._PlotNo}," +
                        $"{item._SizesInEast[0].Text}," +
                        $"{item._SizesInSouth[0].Text}," +
                        $"{item._SizesInWest[0].Text}," +
                        $"{item._SizesInNorth[0].Text}," +
                        String.Format("{0:0.00}", item._PlotArea) + "," +
                        String.Format("{0:0.00}", item._MortgageArea) + "," +
                        String.Format("{0:0.00}", item._AmenityArea) + "," +
                        $"{Convert.ToString(string.Join("|", combinedText.ToArray()))}," +
                        $"{item._EastInfo}," +
                        $"{item._SouthInfo}," +
                        $"{item._WestInfo}," +
                        $"{item._NorthInfo}";

                    sw.WriteLine(textValue1);
                }

                string textValue = $"," +
                       $"," +
                       $"," +
                       $"," +
                       $"," +
                       $"{combinedPlots.Select(x => x._PlotArea).ToArray().Sum():0.00}," +
                       $"{combinedPlots.Select(x => x._MortgageArea).ToArray().Sum():0.00}," +
                       $"{combinedPlots.Select(x => x._AmenityArea).ToArray().Sum():0.00}";

                sw.WriteLine(textValue);

                sw.WriteLine("Total Applicants Site Area ".PadRight(33) + "= " + $"{String.Format("{0:0.00}", RoundLengthValue(SiteInfo.TotalSiteArea))}");

            }
        }

        private List<Polyline> GetPolylinesUsingCrossPolygon(List<Point3d> points, Transaction acTrans, string LayerName)
        {
            List<Polyline> polylines = new List<Polyline>();

            PromptSelectionResult acSSPromptPoly = ed.SelectCrossingWindow(/*new Point3dCollection(SortPoints(points.ToArray()).ToArray())*/ points[0], points[1], CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, LayerName));

            if (acSSPromptPoly.Status == PromptStatus.OK)
            {
                SelectionSet acSSetPoly = acSSPromptPoly.Value;

                foreach (SelectedObject acSSObjPoly in acSSetPoly)
                {
                    if (acSSObjPoly != null)
                    {
                        Polyline acPoly2 = acTrans.GetObject(acSSObjPoly.ObjectId, OpenMode.ForRead) as Polyline;

                        if (acPoly2 != null && acPoly2.Closed)
                        {
                            polylines.Add(acPoly2);
                        }
                    }
                }
            }

            return polylines;
        }

        private double CalculateArea(List<Point3d> points)
        {
            // Sort the points to form a proper closed polyline
            List<Point3d> sortedPoints = SortPoints(points.ToArray());

            // Create a closed polyline from the sorted points
            Polyline polyline = CreateClosedPolyline(sortedPoints);

            // Calculate the area of the polyline
            double area = polyline.Area;

            return area;
        }

        private double CalculateAreaAndCreatePolyline(List<Point3d> points)
        {
            // Sort the points to form a proper closed polyline
            List<Point3d> sortedPoints = SortPoints(points.ToArray());

            // Create a closed polyline from the sorted points
            Polyline polyline = CreateClosedPolylineAndCommit(sortedPoints);

            // Calculate the area of the polyline
            double area = polyline.Area;

            return area;
        }

        private void CreatePoints(List<Point3d> points)
        {
            // Sort the points to form a proper closed polyline
            List<Point3d> sortedPoints = SortPoints(points.ToArray());

            // Create points from the sorted points
            CreatePointsAndCommit(sortedPoints);
        }

        private void CreatePoints(Plot item)
        {
            Point3dCollection intersectionPointsOtherthanBoundary = new Point3dCollection();
            double totalAreainSVNo = 0;
            foreach (SurveyNo svno in item._ParentSurveyNos)
            {
                List<Point3d> intersectionPoints = GetIntersections(item._Polyline, svno._Polyline);
                List<Point3d> uniquePoints = RemoveConsecutiveDuplicates(intersectionPoints);

                //add points to intersectionPointsOtherthanBoundary to draw red color boundary
                foreach (Point3d uniquePoint in uniquePoints)
                {
                    if (!svno._PolylinePoints.Contains(uniquePoint))
                        intersectionPointsOtherthanBoundary.Add(uniquePoint);
                }

                //condition to eliminate 0 areas in some survey no's ex: plot no.65
                //if (item.AreaInSurveyNo[svno] > Constants.minArea)
                //    combinedText.Add($"{svno.DocumentNo + "-" + svno._SurveyNo + "-" + String.Format("{0:0.00}", item.AreaInSurveyNo[svno]) + "-" + svno.LandLordName }");

                totalAreainSVNo += item.AreaInSurveyNo[svno];
            }

            double areaDifference = Math.Abs(totalAreainSVNo - item._Area);
            if (areaDifference >= Constants.areaTolerance)
            {
                if (intersectionPointsOtherthanBoundary.Count > 0)
                {
                    //22.08.2024, validate logic
                    //create polyline code commented as some times we will get only single point, so creating points only
                    //double area = CalculateAreaAndCreatePolyline(intersectionPointsOtherthanBoundary.Cast<Point3d>().ToList());

                    //Create points at the unidentified areas
                    CreatePoints(intersectionPointsOtherthanBoundary.Cast<Point3d>().ToList());
                }
            }
        }

        private Polyline CreateClosedPolyline(List<Point3d> points)
        {
            Polyline polyline = new Polyline();

            // Add points to polyline
            for (int i = 0; i < points.Count; i++)
            {
                Point2d pt2d = new Point2d(points[i].X, points[i].Y);
                polyline.AddVertexAt(i, pt2d, 0, 0, 0);
            }

            polyline.Closed = true; // Close the polyline

            return polyline;
        }

        private Polyline CreateClosedPolylineAndCommit(List<Point3d> points)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database db = acDoc.Database;
            Polyline polyline = new Polyline();

            //Start a transaction
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // Open the BlockTable for read
                BlockTable blockTable = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);

                // Open the LayerTable for read
                LayerTable layerTable = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                string layerName = Constants.FreeSpaceLayer;

                if (layerTable.Has(layerName))
                {
                    //get layers object id
                    ObjectId layerId = layerTable[layerName];

                    //open the layer for read
                    LayerTableRecord layerTableRecord = trans.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;

                    //set the layer as current layer
                    db.Clayer = layerId;
                }

                // Add points to polyline
                for (int i = 0; i < points.Count; i++)
                {
                    Point2d pt2d = new Point2d(points[i].X, points[i].Y);
                    polyline.AddVertexAt(i, pt2d, 0, 0, 0);
                }

                polyline.Closed = true; // Close the polyline

                //Add the polyline to the current space
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                btr.AppendEntity(polyline);
                trans.AddNewlyCreatedDBObject(polyline, true);

                //Commit the transaction
                trans.Commit();
            }

            return polyline;
        }

        private void CreatePointsAndCommit(List<Point3d> points)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database db = acDoc.Database;

            //Start a transaction
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // Open the BlockTable for read
                BlockTable blockTable = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);

                // Open the LayerTable for read
                LayerTable layerTable = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                string layerName = Constants.FreeSpaceLayer;

                if (layerTable.Has(layerName))
                {
                    //get layers object id
                    ObjectId layerId = layerTable[layerName];

                    //open the layer for read
                    LayerTableRecord layerTableRecord = trans.GetObject(layerId, OpenMode.ForRead) as LayerTableRecord;

                    //set the layer as current layer
                    db.Clayer = layerId;
                }

                //// Add points to polyline
                //for (int i = 0; i < points.Count; i++)
                //{
                //    Point2d pt2d = new Point2d(points[i].X, points[i].Y);
                //    polyline.AddVertexAt(i, pt2d, 0, 0, 0);
                //}

                //polyline.Closed = true; // Close the polyline

                //Add the polyline to the current space
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Add points
                for (int i = 0; i < points.Count; i++)
                {
                    Point3d pt3d = new Point3d(points[i].X, points[i].Y, points[i].Z);
                    DBPoint dBPoint = new DBPoint(pt3d);

                    btr.AppendEntity(dBPoint);
                    trans.AddNewlyCreatedDBObject(dBPoint, true);
                }

                //btr.AppendEntity(polyline);
                //trans.AddNewlyCreatedDBObject(polyline, true);

                //Commit the transaction
                trans.Commit();

                //set pdmode system variable
                //Application.SetSystemVariable("PDMODE", 35);

                //set PDSIZE to set point size
                //Application.SetSystemVariable("PDSIZE", 2.0);
            }
        }

        private List<Point3d> SortPoints(Point3d[] points)
        {
            // This is a simple example and may not handle all cases
            // For complex cases, consider using computational geometry libraries
            // For now, we'll assume points are approximately ordered correctly

            // Convert to Point2d for sorting
            List<Point2d> point2dList = points.Select(p => new Point2d(p.X, p.Y)).ToList();

            // Perform a sort based on polar angle from centroid
            Point2d centroid = new Point2d(point2dList.Average(p => p.X), point2dList.Average(p => p.Y));
            point2dList.Sort((p1, p2) => ComparePolarAngles(p1, p2, centroid));

            // Convert back to Point3d
            List<Point3d> sortedPoints = point2dList.Select(p => new Point3d(p.X, p.Y, 0)).ToList();

            return sortedPoints;
        }

        private int ComparePolarAngles(Point2d p1, Point2d p2, Point2d centroid)
        {
            double angle1 = Math.Atan2(p1.Y - centroid.Y, p1.X - centroid.X);
            double angle2 = Math.Atan2(p2.Y - centroid.Y, p2.X - centroid.X);
            return angle1.CompareTo(angle2);
        }

        public double CalculateAngleWithAxis(Point3d startPoint, Point3d endPoint, Vector3d axis)
        {
            // Calculate the direction vector of the line segment
            Vector3d direction = endPoint - startPoint;

            // Normalize the direction and axis vectors
            direction = direction.GetNormal();
            axis = axis.GetNormal();

            // Calculate the dot product
            double dotProduct = direction.DotProduct(axis);

            // Calculate the angle in radians
            double angleRadians = Math.Acos(dotProduct);

            // Convert the angle from radians to degrees
            double angleDegrees = angleRadians * (180.0 / Math.PI);

            return angleDegrees;
        }

        private List<Point3d> GetIntersections(Polyline poly1, Polyline poly2)
        {
            List<Point3d> intersectionPoints = new List<Point3d>();

            using (Transaction acTrans = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction())
            {
                Entity ent1 = (Entity)acTrans.GetObject(poly1.ObjectId, OpenMode.ForRead);
                Entity ent2 = (Entity)acTrans.GetObject(poly2.ObjectId, OpenMode.ForRead);

                Point3dCollection points = new Point3dCollection();
                ent1.IntersectWith(ent2, Intersect.OnBothOperands, points, IntPtr.Zero, IntPtr.Zero);

                foreach (Point3d pt in points)
                {
                    intersectionPoints.Add(pt);
                }

                acTrans.Commit();
            }

            return intersectionPoints;
        }

        private List<Point3d> RemoveConsecutiveDuplicates(List<Point3d> points)
        {
            if (points == null || points.Count == 0) return points;

            List<Point3d> uniquePoints = new List<Point3d> { points[0] };

            for (int i = 1; i < points.Count; i++)
            {
                if (!points[i].IsEqualTo(points[i - 1], new Tolerance(1e-4, 1e-4)))
                {
                    uniquePoints.Add(points[i]);
                }
            }

            return uniquePoints;
        }

        private Point3d getCenter(Polyline polyline)
        {
            Extents3d polyExtents = polyline.GeometricExtents;
            // Calculate the center point of the polyline
            Point3d centerPoint = new Point3d(
                (polyExtents.MinPoint.X + polyExtents.MaxPoint.X) / 2,
                (polyExtents.MinPoint.Y + polyExtents.MaxPoint.Y) / 2,
                (polyExtents.MinPoint.Z + polyExtents.MaxPoint.Z) / 2);

            return centerPoint;
        }

        private void FillLineSegmentsAndPointsByDirection(Polyline polyline, Plot myObject)
        {
            #region OldLogic

            //Iterate through the polyline segments
            //for (int i = 0; i < polyline.NumberOfVertices; i++)
            //{
            //    myObject._PolylinePoints.Add(polyline.GetPoint3dAt(i));

            //    if (polyline.GetSegmentType(i) == SegmentType.Line)
            //    {
            //        LineSegment2d segment = polyline.GetLineSegment2dAt(i);

            //        // Determine the direction of the segment
            //        if (segment.StartPoint.X == segment.EndPoint.X)
            //        {
            //            // Vertical line
            //            if (segment.StartPoint.Y < segment.EndPoint.Y)
            //            {
            //                myObject.northPoints.Add(segment.StartPoint);
            //                myObject.northPoints.Add(segment.EndPoint);
            //            }
            //            else
            //            {
            //                myObject.southPoints.Add(segment.StartPoint);
            //                myObject.southPoints.Add(segment.EndPoint);
            //            }
            //        }
            //        else if (segment.StartPoint.Y == segment.EndPoint.Y)
            //        {
            //            // Horizontal line
            //            if (segment.StartPoint.X < segment.EndPoint.X)
            //            {
            //                myObject.eastPoints.Add(segment.StartPoint);
            //                myObject.eastPoints.Add(segment.EndPoint);
            //            }
            //            else
            //            {
            //                myObject.westPoints.Add(segment.StartPoint);
            //                myObject.westPoints.Add(segment.EndPoint);
            //            }
            //        }
            //    }
            //}

            #endregion

            //ToDo - need to add multiple line segments if line is split into multiple segments to handle multiple dimensions (Need to include
            //angle logic to identify correct direction of line segment ex: x > +ve & Y > +ve can be in east or north 

            //Fill all line segments in all directions
            for (int i = 0; i < polyline.NumberOfVertices; i++)
            {
                //if (polyline.GetSegmentType(i) == SegmentType.Line)
                if (polyline.GetLineSegmentAt(i).Length > 0.5)
                {
                    Point3d vertex1 = polyline.GetLineSegmentAt(i).StartPoint;
                    Point3d vertex2 = polyline.GetLineSegmentAt(i).EndPoint;

                    Vector3d direction1 = vertex1 - myObject.Center;
                    Vector3d direction2 = vertex2 - myObject.Center;

                    if ((direction1.X < 0 && direction1.Y > 0) && (direction2.X > 0 && direction2.Y > 0) ||
                        (direction2.X < 0 && direction2.Y > 0) && (direction1.X > 0 && direction1.Y > 0))
                    //|| (direction1.X < 0 && direction1.Y > 0) && (direction2.X < 0 && direction2.Y > 0) ||
                    //(direction1.X > 0 && direction1.Y > 0) && (direction2.X > 0 && direction2.Y > 0))
                    {
                        myObject.northLineSegment.Add(polyline.GetLineSegmentAt(i));
                        myObject.northPoints.Add(vertex1);
                        myObject.northPoints.Add(vertex2);
                    }

                    if ((direction1.X < 0 && direction1.Y < 0) && (direction2.X > 0 && direction2.Y < 0) ||
                        (direction2.X < 0 && direction2.Y < 0) && (direction1.X > 0 && direction1.Y < 0))
                    //    ||
                    //    (direction1.X < 0 && direction1.Y < 0) && (direction2.X < 0 && direction2.Y < 0) ||
                    //    (direction1.X > 0 && direction1.Y < 0) && (direction2.X > 0 && direction2.Y < 0))
                    {
                        myObject.southLineSegment.Add(polyline.GetLineSegmentAt(i));
                        myObject.southPoints.Add(vertex1);
                        myObject.southPoints.Add(vertex2);
                    }

                    if ((direction1.X > 0 && direction1.Y > 0) && (direction2.X > 0 && direction2.Y < 0) ||
                        (direction2.X > 0 && direction2.Y > 0) && (direction1.X > 0 && direction1.Y < 0))
                    //||
                    //(direction1.X > 0 && direction1.Y > 0) && (direction2.X > 0 && direction2.Y > 0) ||
                    //(direction1.X > 0 && direction1.Y < 0) && (direction2.X > 0 && direction2.Y < 0))
                    {
                        //double angle = CalculateAngleWithAxis(vertex1, vertex2, Vector3d.XAxis);
                        //System.Diagnostics.Debug.Print("East Start\n");
                        //System.Diagnostics.Debug.Print("East " + angle + "\n");

                        //if (angle >= 90 && angle <= 180)
                        //{

                        //}
                        myObject.eastLineSegment.Add(polyline.GetLineSegmentAt(i));
                        myObject.eastPoints.Add(vertex1);
                        myObject.eastPoints.Add(vertex2);
                    }

                    if ((direction1.X < 0 && direction1.Y > 0) && (direction2.X < 0 && direction2.Y < 0) ||
                        (direction2.X < 0 && direction2.Y > 0) && (direction1.X < 0 && direction1.Y < 0))
                    //||
                    //(direction1.X < 0 && direction1.Y > 0) && (direction2.X < 0 && direction2.Y > 0) ||
                    //(direction1.X < 0 && direction1.Y > 0) && (direction2.X < 0 && direction2.Y < 0))
                    {
                        myObject.westLineSegment.Add(polyline.GetLineSegmentAt(i));
                        myObject.westPoints.Add(vertex1);
                        myObject.westPoints.Add(vertex2);
                    }

                    #region Old Logic to collect all direction points
                    //Point3d vertex = polyline.GetPoint3dAt(i);
                    //Vector3d direction = vertex - myObject.Center;
                    //List<Point3d> points = new List<Point3d>() { vertex1, vertex2 };

                    //Looping through start and end points of line

                    //foreach (Point3d point in points)
                    //{
                    //    Vector3d direction = point - myObject.Center;

                    //    if (direction.Y > 0)
                    //    {
                    //        myObject.northPoints.Add(point);
                    //    }
                    //    else if (direction.Y < 0)
                    //    {
                    //        myObject.southPoints.Add(point);
                    //    }

                    //    if (direction.X > 0)
                    //    {
                    //        myObject.eastPoints.Add(point);
                    //    }
                    //    else if (direction.X < 0)
                    //    {
                    //        myObject.westPoints.Add(point);
                    //    }
                    //}
                    #endregion
                }
            }

            #region Old Logic to filter all direction points

            //// Convert the dynamic lists to List<Point3d> and then process them
            //List<Point3d> northPoints = ((IEnumerable<Point3d>)myObject.northPoints).OrderBy(p => p.DistanceTo((Point3d)myObject.Center)).Take(2).ToList();
            //List<Point3d> southPoints = ((IEnumerable<Point3d>)myObject.southPoints).OrderBy(p => p.DistanceTo((Point3d)myObject.Center)).Take(2).ToList();
            //List<Point3d> eastPoints = ((IEnumerable<Point3d>)myObject.eastPoints).OrderBy(p => p.DistanceTo((Point3d)myObject.Center)).Take(2).ToList();
            //List<Point3d> westPoints = ((IEnumerable<Point3d>)myObject.westPoints).OrderBy(p => p.DistanceTo((Point3d)myObject.Center)).Take(2).ToList();

            //// Assign the processed lists back to the dynamic object
            //myObject.northPoints = northPoints;
            //myObject.southPoints = southPoints;
            //myObject.eastPoints = eastPoints;
            //myObject.westPoints = westPoints;

            #endregion
        }

        private SelectionFilter CreateSelectionFilterByStartTypeAndLayer(string startType, string Layer)
        {
            TypedValue[] Filter = { new TypedValue((int)DxfCode.Start, startType),
                                new TypedValue((int)DxfCode.LayerName, Layer)
                              };

            SelectionFilter acSelFtr = new SelectionFilter(Filter);
            return acSelFtr;
        }

        private SelectionFilter CreateSelectionFilterByStartType(string startType)
        {
            TypedValue[] Filter = { new TypedValue((int)DxfCode.Start, startType),
                              };

            SelectionFilter acSelFtr = new SelectionFilter(Filter);
            return acSelFtr;
        }

        private SelectionFilter CreateSelectionFilterByStartTypeAndLayer(string startType1, string startType2, string Layer)
        {
            TypedValue[] Filter = { new TypedValue((int)DxfCode.Start, startType1),
                                new TypedValue((int)DxfCode.Start, startType2),
                                new TypedValue((int)DxfCode.LayerName, Layer)
                              };

            SelectionFilter acSelFtr = new SelectionFilter(Filter);
            return acSelFtr;
        }

        private Plot FillPlotObject(List<SurveyNo> surveyNos, Polyline acPoly2, Plot plotNo, SurveyNo surveyNo, Transaction acTrans, string TextLayerName/*, string dimensionLayerName*/)
        {
            // new logic added
            //var existingPlotNos = surveyNos.SelectMany(x => x._PlotNos).Where
            //(x => x._Polyline.ObjectId == acPoly2.ObjectId).ToList();

            //if (existingPlotNos.Count > 0)
            //{
            //    plotNo = existingPlotNos[0];
            //    plotNo._ParentSurveyNos.Add(surveyNo);
            //}

            //else
            //{
            plotNo._Polyline = acPoly2; //assign polyline
            plotNo.Center = getCenter(plotNo._Polyline);

            //Collect all Points
            for (int i = 0; i < acPoly2.NumberOfVertices; i++)
            {
                plotNo._PolylinePoints.Add(acPoly2.GetPoint3dAt(i));
            }

            //fill plot number text, area and parent survey No
            string plotNoText = GetTextFromLayer(acTrans, plotNo._PolylinePoints, Constants.TEXT, TextLayerName);
            if (!string.IsNullOrEmpty(plotNoText))
            {
                plotNo._PlotNo = plotNoText;//assign plotNo Text
                plotNo._Area = plotNo._Polyline.Area; /*Math.Round(plotNo._Polyline.Area, Constants.AreaDecimals);*/
                plotNo._ParentSurveyNos.Add(surveyNo);
            }
            else
            {
                plotNo._PlotNo = GetTextFromLayer(acTrans, plotNo._PolylinePoints, Constants.MTEXT, TextLayerName);//assign plotNo Text
                plotNo._Area = plotNo._Polyline.Area;/* Math.Round(plotNo._Polyline.Area, Constants.AreaDecimals);*/
                plotNo._ParentSurveyNos.Add(surveyNo);
            }

            #region Old Code
            //fill plot number text assuming text is of single text, area and parent survey No
            //PromptSelectionResult textSelResult = ed.SelectWindowPolygon(plotNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer(Constants.TEXT, TextLayerName));

            //if (textSelResult.Status == PromptStatus.OK)
            //{
            //    DBText textEntity = acTrans.GetObject(textSelResult.Value[0].ObjectId, OpenMode.ForRead) as DBText;
            //    if (textEntity != null)
            //    {
            //        string fval2 = textEntity.TextString;
            //        plotNo._PlotNo = fval2;
            //        plotNo._Area = Math.Round(plotNo._Polyline.Area, 3);
            //        plotNo._ParentSurveyNos.Add(surveyNo);
            //    }
            //}

            ////fill plot number text assuming text is of multiline text, area and parent survey No
            //PromptSelectionResult textSelResult1 = ed.SelectWindowPolygon(plotNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer(Constants.MTEXT, TextLayerName));

            //if (textSelResult1.Status == PromptStatus.OK)
            //{
            //    MText textEntity = acTrans.GetObject(textSelResult1.Value[0].ObjectId, OpenMode.ForRead) as MText;
            //    if (textEntity != null)
            //    {
            //        string fval2 = textEntity.Text;
            //        plotNo._PlotNo = fval2;
            //        plotNo._Area = Math.Round(plotNo._Polyline.Area, 3);
            //        plotNo._ParentSurveyNos.Add(surveyNo);
            //    }
            //}
            #endregion

            //Tested after commenting this method, get all dimensions of plot and fill in plot object
            //PromptSelectionResult dimSelResult = ed.SelectCrossingPolygon(plotNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer("DIMENSION", dimensionLayerName));

            //if (dimSelResult.Status == PromptStatus.OK)
            //{
            //    SelectionSet dimSelSet = dimSelResult.Value;
            //    foreach (SelectedObject acSSObj in dimSelSet)
            //    {
            //        if (acSSObj != null)
            //        {
            //            Dimension dimEntity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Dimension;
            //            if (dimEntity != null)
            //            {
            //                string dimValue = dimEntity.DimensionText;

            //                if (Convert.ToDouble(dimValue) > 0.5)
            //                    plotNo._AllDims.Add(new SDimension(dimEntity, dimEntity.DimensionText, dimEntity.TextPosition));
            //            }
            //        }
            //    }
            //}

            //fill sizes by direction
            FillLineSegmentsAndPointsByDirection(plotNo._Polyline, plotNo);

            //Tested after commenting this method - main method responsible for sorting dimensions
            //FillSizesByDirection(plotNo);

            //New method added to get sizes from length itself no need to go for dimension layers
            FillSizesByLength(plotNo);


            //}

            return plotNo;

            //surveyNo._PlotNos.Add(plotNo); //add plotNo to SurveyNo List
        }

        private void FillMortgageObject(List<SurveyNo> surveyNos, Polyline acPoly2, Mortgage plotNo, SurveyNo surveyNo, Transaction acTrans)
        {
            // new logic added
            var existingMortgagePlotNos = surveyNos.SelectMany(x => x._MortgagePlotNos).Where
            (x => x._Polyline.ObjectId == acPoly2.ObjectId).ToList();

            if (existingMortgagePlotNos.Count > 0)
            {
                plotNo = existingMortgagePlotNos[0];
                plotNo._ParentSurveyNos.Add(surveyNo);
                //FillSizesByDirection(plotNo, acTrans);
            }

            else
            {
                plotNo._Polyline = acPoly2; //assign polyline
                plotNo.Center = getCenter(plotNo._Polyline);
                //FillAllPointsAndByDirection(plotNo._Polyline, plotNo);
            }

            surveyNo._MortgagePlotNos.Add(plotNo); //add plotNo to SurveyNo List
        }

        private void FillSizesByDirection(Plot plotNo)
        {
            //for your information, filling sizes by direction based on available dimensions
            //-> you can use polyline.length also instead of dimension
            //-> you can change dimension type to text if text is placed for dimension

            for (int i = 0; i < plotNo._AllDims.Count; i++)
            {
                Point3d vertex = plotNo._AllDims[i].position;

                Vector3d direction = vertex - plotNo.Center;

                if (direction.Y > 0)
                {
                    plotNo._SizesInNorth.Add(plotNo._AllDims[i]);
                }
                else if (direction.Y < 0)
                {
                    plotNo._SizesInSouth.Add(plotNo._AllDims[i]);
                }

                if (direction.X > 0)
                {
                    plotNo._SizesInEast.Add(plotNo._AllDims[i]);
                }
                else if (direction.X < 0)
                {
                    plotNo._SizesInWest.Add(plotNo._AllDims[i]);
                }
            }

            //ToDo - need to add multiple dimensions if line is split into multiple segments

            plotNo._SizesInNorth = plotNo._SizesInNorth.OrderBy(p => p.position.DistanceTo(plotNo.northLineSegment[0].MidPoint)).Take(2).ToList();
            plotNo._SizesInSouth = plotNo._SizesInSouth.OrderBy(p => p.position.DistanceTo(plotNo.southLineSegment[0].MidPoint)).Take(2).ToList();
            plotNo._SizesInEast = plotNo._SizesInEast.OrderBy(p => p.position.DistanceTo(plotNo.eastLineSegment[0].MidPoint)).Take(2).ToList();
            plotNo._SizesInWest = plotNo._SizesInWest.OrderBy(p => p.position.DistanceTo(plotNo.westLineSegment[0].MidPoint)).Take(2).ToList();

            //plotNo._SizesInNorth = plotNo._SizesInNorth.OrderByDescending(x => x.position.Y).Take(3).ToList().OrderBy(p => p.position.DistanceTo(plotNo.Center)).Take(2).ToList();

            //plotNo._SizesInNorth = plotNo._SizesInNorth.OrderBy(p => p.position.DistanceTo(plotNo.Center)).Take(2).ToList();
            //plotNo._SizesInSouth = plotNo._SizesInSouth.OrderBy(p => p.position.DistanceTo(plotNo.Center)).Take(2).ToList();
            //plotNo._SizesInEast = plotNo._SizesInEast.OrderBy(p => p.position.DistanceTo(plotNo.Center)).Take(2).ToList();
            //plotNo._SizesInWest = plotNo._SizesInWest.OrderBy(p => p.position.DistanceTo(plotNo.Center)).Take(2).ToList();

        }

        private void FillSizesByLength(Plot plotNo)
        {
            try
            {
                plotNo._SizesInNorth.Add(new SDimension(null, String.Format("{0:0.00}", RoundLengthValue(plotNo.northLineSegment[0].Length)), new Point3d(0, 0, 0)));
                plotNo._SizesInSouth.Add(new SDimension(null, String.Format("{0:0.00}", RoundLengthValue(plotNo.southLineSegment[0].Length)), new Point3d(0, 0, 0)));
                plotNo._SizesInEast.Add(new SDimension(null, String.Format("{0:0.00}", RoundLengthValue(plotNo.eastLineSegment[0].Length)), new Point3d(0, 0, 0)));
                plotNo._SizesInWest.Add(new SDimension(null, String.Format("{0:0.00}", RoundLengthValue(plotNo.westLineSegment[0].Length)), new Point3d(0, 0, 0)));
            }
            catch
            {
                //error in filling sizes by polyline length
            }
        }

        private double RoundLengthValue(double value)
        {
            return Math.Round(value, 2);
        }

        private void FillSizesByDirection(Plot plotNo, Transaction acTrans)
        {
            //plotNo._SizeInEast = GetSize(plotNo, acTrans, new Point3dCollection(plotNo.eastPoints.ToArray()));
            //plotNo._SizeInSouth = GetSize(plotNo, acTrans, new Point3dCollection(plotNo.southPoints.ToArray()));
            //plotNo._SizeInWest = GetSize(plotNo, acTrans, new Point3dCollection(plotNo.westPoints.ToArray()));
            //plotNo._SizeInNorth = GetSize(plotNo, acTrans, new Point3dCollection(plotNo.northPoints.ToArray()));
        }

        //private string GetSize(Plot plotNo, Transaction acTrans, Point3dCollection point3dCollection)
        //{
        //    PromptSelectionResult textSelResult = ed.SelectCrossingPolygon(point3dCollection, CreateSelectionFilterByStartTypeAndLayer("DIMENSION", Constants.IndivPlotDimLayer));

        //    if (textSelResult.Status == PromptStatus.OK)
        //    {
        //        Dimension textEntity = acTrans.GetObject(textSelResult.Value[0].ObjectId, OpenMode.ForRead) as Dimension;
        //        if (textEntity != null)
        //        {
        //            string fval2 = textEntity.DimensionText;
        //            return fval2;
        //        }
        //    }

        //    return "";
        //}

        private string GetTextFromLayer(Transaction acTrans, Point3dCollection point3dCollection, string textType, string layerName)
        {
            PromptSelectionResult acSSPromptText = ed.SelectCrossingPolygon(point3dCollection, CreateSelectionFilterByStartTypeAndLayer(textType, layerName));

            if (acSSPromptText.Status == PromptStatus.OK)
            {
                SelectionSet acSSetText = acSSPromptText.Value;
                string val2;
                try
                {
                    DBText acText = acTrans.GetObject(acSSetText[0].ObjectId, OpenMode.ForRead) as DBText;
                    val2 = acText.TextString;
                }

                catch
                {
                    MText acText = acTrans.GetObject(acSSetText[0].ObjectId, OpenMode.ForRead) as MText;
                    val2 = acText.Text;
                }

                return val2;
            }

            return "";
        }

        private List<string> GetListTextFromLayer(Transaction acTrans, Point3dCollection point3dCollection, string textType, string layerName)
        {
            List<string> collection = new List<string>();

            PromptSelectionResult acSSPromptText = ed.SelectCrossingPolygon(point3dCollection, CreateSelectionFilterByStartTypeAndLayer(textType, layerName));

            if (acSSPromptText.Status == PromptStatus.OK)
            {
                SelectionSet acSSetTexts = acSSPromptText.Value;

                foreach (SelectedObject acSSetText in acSSetTexts)
                {
                    if (acSSetText != null)
                    {
                        string val2;
                        try
                        {
                            DBText acText = acTrans.GetObject(acSSetText.ObjectId, OpenMode.ForRead) as DBText;
                            val2 = acText.TextString;
                        }

                        catch
                        {
                            MText acText = acTrans.GetObject(acSSetText.ObjectId, OpenMode.ForRead) as MText;
                            val2 = acText.Text;
                        }

                        collection.Add(val2); //add to list
                    }
                }
            }

            return collection;
        }

        private List<(ObjectId, string)> GetListTextAndObjectFromLayer(Transaction acTrans, Point3dCollection point3dCollection, string textType, string layerName)
        {
            List<(ObjectId, string)> collection = new List<(ObjectId, string)>();

            PromptSelectionResult acSSPromptText = ed.SelectCrossingPolygon(point3dCollection, CreateSelectionFilterByStartTypeAndLayer(textType, layerName));

            if (acSSPromptText.Status == PromptStatus.OK)
            {
                SelectionSet acSSetTexts = acSSPromptText.Value;

                foreach (SelectedObject acSSetText in acSSetTexts)
                {
                    if (acSSetText != null)
                    {
                        string val2;
                        try
                        {
                            DBText acText = acTrans.GetObject(acSSetText.ObjectId, OpenMode.ForRead) as DBText;
                            val2 = acText.TextString;
                        }

                        catch
                        {
                            MText acText = acTrans.GetObject(acSSetText.ObjectId, OpenMode.ForRead) as MText;
                            val2 = acText.Text;
                        }
                        collection.Add((acSSetText.ObjectId, val2)); //add to list
                    }
                }
            }

            return collection;
        }

        public bool IsPointInsidePolyline(Point3d point, Polyline polyline)
        {
            // Convert Point3d to Point2d
            Point2d pt2d = new Point2d(point.X, point.Y);

            // Ensure the polyline is closed
            if (!polyline.Closed)
                return false;

            // Convert polyline vertices to Point2d
            List<Point2d> vertices = new List<Point2d>();
            for (int i = 0; i < polyline.NumberOfVertices; i++)
            {
                Point2d vertex = polyline.GetPoint2dAt(i);
                vertices.Add(vertex);
            }

            // Perform the ray-casting algorithm
            return IsPointInPolygon(pt2d, vertices);
        }

        private bool IsPointInPolygon(Point2d point, List<Point2d> vertices)
        {
            int n = vertices.Count;
            bool inside = false;

            double px = point.X;
            double py = point.Y;

            for (int i = 0, j = n - 1; i < n; j = i++)
            {
                double viX = vertices[i].X;
                double viY = vertices[i].Y;
                double vjX = vertices[j].X;
                double vjY = vertices[j].Y;

                bool intersect = ((viY > py) != (vjY > py)) &&
                                 (px < (vjX - viX) * (py - viY) / (vjY - viY) + viX);
                if (intersect)
                    inside = !inside;
            }

            return inside;
        }

        public double GetAreaByLayer(string layerName, Transaction acTrans)
        {
            double TotalArea = 0;

            PromptSelectionResult acSSPrompt1 = ed.SelectAll(CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, layerName));

            if (acSSPrompt1.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = acSSPrompt1.Value;

                foreach (SelectedObject acSSObj in acSSet)
                {
                    if (acSSObj != null)
                    {
                        Polyline acPoly = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;

                        if (acPoly != null && acPoly.Closed)
                            TotalArea += acPoly.Area;
                    }
                }
            }

            return TotalArea;
        }

        public List<(ObjectId, string)> GetListTextAndObjectIdFromLayer(string layerName, Transaction acTrans)
        {
            List<(ObjectId, string)> textArray = new List<(ObjectId, string)>();

            PromptSelectionResult acSSPrompt1 = ed.SelectAll(CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, layerName));

            if (acSSPrompt1.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = acSSPrompt1.Value;

                foreach (SelectedObject acSSObj in acSSet)
                {
                    if (acSSObj != null)
                    {
                        Polyline acPoly = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;

                        if (acPoly != null && acPoly.Closed)
                        {
                            //Collect all Points
                            Point3dCollection point3DCollection = new Point3dCollection();

                            for (int i = 0; i < acPoly.NumberOfVertices; i++)
                            {
                                point3DCollection.Add(acPoly.GetPoint3dAt(i));
                            }

                            textArray.AddRange(GetListTextAndObjectFromLayer(acTrans, point3DCollection, Constants.TEXT, layerName));
                            textArray.AddRange(GetListTextAndObjectFromLayer(acTrans, point3DCollection, Constants.MTEXT, layerName));
                        }
                    }
                }
            }

            return textArray;
        }

        public List<string> GetListTextFromLayer(string layerName, Transaction acTrans)
        {
            List<string> textArray = new List<string>();

            PromptSelectionResult acSSPrompt1 = ed.SelectAll(CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, layerName));

            if (acSSPrompt1.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = acSSPrompt1.Value;

                foreach (SelectedObject acSSObj in acSSet)
                {
                    if (acSSObj != null)
                    {
                        Polyline acPoly = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;

                        if (acPoly != null && acPoly.Closed)
                        {
                            //Collect all Points
                            Point3dCollection point3DCollection = new Point3dCollection();

                            for (int i = 0; i < acPoly.NumberOfVertices; i++)
                            {
                                point3DCollection.Add(acPoly.GetPoint3dAt(i));
                            }

                            textArray.AddRange(GetListTextFromLayer(acTrans, point3DCollection, Constants.TEXT, layerName));
                            textArray.AddRange(GetListTextFromLayer(acTrans, point3DCollection, Constants.MTEXT, layerName));
                        }
                    }
                }
            }

            return textArray;
        }

        public List<ObjectId> GetPolyLines(string layerName, Transaction acTrans)
        {
            List<ObjectId> polylinesList = new List<ObjectId>();

            PromptSelectionResult acSSPrompt1 = ed.SelectAll(CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, layerName));

            if (acSSPrompt1.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = acSSPrompt1.Value;

                foreach (SelectedObject acSSObj in acSSet)
                {
                    if (acSSObj != null)
                    {
                        Polyline acPoly = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;
                        polylinesList.Add(acPoly.ObjectId);
                    }
                }
            }

            return polylinesList;
        }

        public Dictionary<ObjectId, string> GetListTextDictionaryFromLayer(string layerName, Transaction acTrans, List<(ObjectId, List<string>)> polylineIdsWithMultipleSurveyNos)
        {
            Dictionary<ObjectId, string> textArray = new Dictionary<ObjectId, string>();

            PromptSelectionResult acSSPrompt1 = ed.SelectAll(CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, layerName));

            if (acSSPrompt1.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = acSSPrompt1.Value;

                foreach (SelectedObject acSSObj in acSSet)
                {
                    if (acSSObj != null)
                    {
                        Polyline acPoly = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;

                        if (acPoly != null && acPoly.Closed)
                        {
                            //Collect all Points
                            Point3dCollection point3DCollection = new Point3dCollection();

                            for (int i = 0; i < acPoly.NumberOfVertices; i++)
                            {
                                point3DCollection.Add(acPoly.GetPoint3dAt(i));
                            }

                            List<string> textValues = GetListTextFromLayer(acTrans, point3DCollection, Constants.TEXT, layerName);

                            if (textValues.Count > 1)
                            {
                                polylineIdsWithMultipleSurveyNos.Add((acPoly.ObjectId, textValues));
                            }
                            if (textValues.Count == 1)
                            {
                                if (!string.IsNullOrEmpty(textValues[0]))
                                    textArray.Add(acPoly.ObjectId, textValues[0]);
                            }

                            List<string> mTextValues = GetListTextFromLayer(acTrans, point3DCollection, Constants.MTEXT, layerName);

                            if (mTextValues.Count > 1)
                            {
                                polylineIdsWithMultipleSurveyNos.Add((acPoly.ObjectId, mTextValues));
                            }
                            if (mTextValues.Count == 1)
                            {
                                if (!string.IsNullOrEmpty(mTextValues[0]))
                                    textArray.Add(acPoly.ObjectId, mTextValues[0]);
                            }
                        }
                    }
                }
            }

            return textArray;
        }

        //private string GetTextFromLayer2(Transaction acTrans, Point3dCollection point3dCollection, string textType, string layerName)
        //{
        //    PromptSelectionResult acSSPromptText = ed.SelectCrossingPolygon(point3dCollection, CreateSelectionFilterByStartTypeAndLayer(textType, layerName));

        //    if (acSSPromptText.Status == PromptStatus.OK)
        //    {
        //        SelectionSet acSSetText = acSSPromptText.Value;
        //        MText acText = acTrans.GetObject(acSSetText[0].ObjectId, OpenMode.ForRead) as MText;
        //        string val2 = acText.Text;
        //        return val2;
        //    }

        //    return "";
        //}


        #region Check whether any polyline is outside the plot layer

        //public void CheckPolylineContainment()
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;

        //    using (Transaction tr = db.TransactionManager.StartTransaction())
        //    {
        //        // Assuming we have the ObjectIds of the two polylines
        //        ObjectId outerPolylineId = new ObjectId(); // Set this to the ObjectId of the outer polyline
        //        ObjectId innerPolylineId = new ObjectId(); // Set this to the ObjectId of the inner polyline

        //        Polyline outerPolyline = tr.GetObject(outerPolylineId, OpenMode.ForRead) as Polyline;
        //        Polyline innerPolyline = tr.GetObject(innerPolylineId, OpenMode.ForRead) as Polyline;

        //        if (outerPolyline != null && innerPolyline != null)
        //        {
        //            bool isInside = IsPolylineCompletelyInside(outerPolyline, innerPolyline);

        //            if (isInside)
        //            {
        //                doc.Editor.WriteMessage("\nThe inner polyline is completely inside the outer polyline.");
        //            }
        //            else
        //            {
        //                doc.Editor.WriteMessage("\nThe inner polyline is not completely inside the outer polyline.");
        //            }
        //        }

        //        tr.Commit();
        //    }
        //}

        //private bool IsPolylineCompletelyInside(Polyline outerPolyline, Polyline innerPolyline)
        //{
        //    // 1. Check if all vertices of the inner polyline are inside or on the boundary of the outer polyline
        //    for (int i = 0; i < innerPolyline.NumberOfVertices; i++)
        //    {
        //        Point3d vertex = innerPolyline.GetPoint3dAt(i);

        //        if (!IsPointInsideOrOnPolyline(outerPolyline, vertex))
        //        {
        //            return false; // If any vertex is outside, the polyline is not completely inside
        //        }
        //    }

        //    // 2. Check for intersection between the inner and outer polylines
        //    if (CheckPolylinesIntersection(outerPolyline, innerPolyline))
        //    {
        //        return false; // If there is any intersection, the inner polyline is not completely inside
        //    }

        //    return true; // All vertices are inside/on and there's no intersection
        //}

        //private bool IsPointInsideOrOnPolyline(Polyline polyline, Point3d point)
        //{
        //    // Convert polyline to Region to check point containment
        //    using (var db = polyline.Database)
        //    {
        //        using (var region = new Region())
        //        {
        //            DBObjectCollection dbObjCollection = new DBObjectCollection();
        //            polyline.Explode(dbObjCollection);

        //            DBObjectCollection regionColl = new DBObjectCollection();
        //            Region.CreateFromCurves(dbObjCollection, regionColl);

        //            Region regionEntity = regionColl[0] as Region;

        //            if (regionEntity != null)
        //            {
        //                // Check if point is inside or on the boundary
        //                PointContainment containment = regionEntity.PointContainment(new Point2d(point.X, point.Y));
        //                return containment == PointContainment.Inside || containment == PointContainment.OnBoundary;
        //            }
        //        }
        //    }
        //    return false;
        //}

        //private bool CheckPolylinesIntersection(Polyline polyline1, Polyline polyline2)
        //{
        //    for (int i = 0; i < polyline1.NumberOfVertices; i++)
        //    {
        //        LineSegment3d segment1 = polyline1.GetLineSegmentAt(i);

        //        for (int j = 0; j < polyline2.NumberOfVertices; j++)
        //        {
        //            LineSegment3d segment2 = polyline2.GetLineSegmentAt(j);

        //            if (segment1.IntersectWith(segment2, Intersect.OnBothOperands).Count > 0)
        //            {
        //                return true; // If any segments intersect, the polylines intersect
        //            }
        //        }
        //    }
        //    return false; // No intersection found
        //}

        #endregion


    }
}

//Fill Mortgages

//PromptSelectionResult mortgageSelResult = ed.SelectCrossingPolygon(plotNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer(Constants.LWPOLYLINE, Constants.MortgageLayer));

//if (mortgageSelResult.Status == PromptStatus.OK)
//{
//    SelectionSet dimSelSet = mortgageSelResult.Value;
//    foreach (SelectedObject acSSObj in dimSelSet)
//    {
//        if (acSSObj != null)
//        {
//            Polyline MortEntity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;
//            if (MortEntity != null)
//            {
//                List<Point3d> intersectionPoints = GetIntersections(MortEntity, acPoly2);
//                List<Point3d> uniquePoints = RemoveConsecutiveDuplicates(intersectionPoints);

//                if (uniquePoints.Count > 1)
//                {
//                    plotNo.MortgagePoints.AddRange(uniquePoints);
//                    plotNo.IsMortgageArea = true;

//                    //for (int i = 0; i < MortEntity.NumberOfVertices; i++)
//                    //{
//                    //    plotNo.MortgagePoints.Add(MortEntity.GetPoint3dAt(i));
//                    //}

//                    //    string fval2 = textEntity.TextString;
//                    //plotNo._PlotNo = fval2;
//                    //plotNo._ParentSurveyNos.Add(surveyNo);
//                    //FillSizesByDirection(plotNo, acTrans);
//                    //
//                }
//            }
//        }
//    }
//}
