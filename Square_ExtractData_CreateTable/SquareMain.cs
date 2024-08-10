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

public class MyCommands
{
    private static Editor ed = null;
    List<string> mortgagePlots = new List<string>();
    private static int AreaDecimals = 2;
    private static int uniquePointsIdentifier = 1;
    private static double minArea = 1.0;

    private static double TotalSiteArea = 0;
    private static double PlotsArea = 0;
    //private static double MortgageArea = 0;
    private static double AmenitiesArea = 0;
    private static double OpenSpaceArea = 0;
    private static double UtilityArea = 0;
    private static double InternalRoadsArea = 0;
    private static double SplayArea = 0;
    private static double LeftOverOwnerLandArea = 0;
    private static double RoadWideningArea = 0;
    private static double GreenArea = 0;

    private static double VerifiedArea = 0;
    private static double differenceArea = 0;



    //public Form1 frm;
    //private Thread cadThread1;
    //public MyWPF myWpfForm;
    //public static ViewModelObject viewModel;

    [CommandMethod("PB")]
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

    [CommandMethod("CLAYER")]
    public void CheckAndCreateLayers()
    {
        // Get the current database and transaction manager

        if (IsLicenseExpired())
            return;

        Document acDoc = Application.DocumentManager.MdiActiveDocument;
        Database acCurDb = acDoc.Database;
        Editor ed = acDoc.Editor;

        // Zoom extents
        ed.Command("_.zoom", "_e");

        List<string> layersList = new List<string>()
        {
            Constants.SurveyNoLayer,
            Constants.IndivPlotLayer,
            Constants.IndivPlotDimLayer,
            Constants.MortgageLayer,
            Constants.AmenityLayer,
            Constants.AmenityDimLayer,
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
            Constants.GreenBufferZoneLayer
        };

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

        Application.SetSystemVariable("CMDECHO", 0);

        // Execute layer management commands after the transaction is committed
        foreach (var layerName in layersList)
        {
            ed.Command("_-layer", "t", layerName, "ON", layerName, "U", layerName, "");
        }

        Application.SetSystemVariable("CMDECHO", 1);
    }


    [CommandMethod("MEE")]
    //[CommandMethod("MExport")]

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

        List<string> layersList = new List<string>() { Constants.SurveyNoLayer, Constants.IndivPlotLayer, Constants.IndivPlotDimLayer, Constants.MortgageLayer, Constants.AmenityLayer, Constants.AmenityDimLayer, Constants.DocNoLayer, Constants.LandLordLayer, Constants.InternalRoadLayer, Constants.PlotLayer, Constants.OpenSpaceLayer, Constants.UtilityLayer, Constants.LeftOverOwnerLandLayer, Constants.SideBoundaryLayer, Constants.MainRoadLayer, Constants.SplayLayer, Constants.RoadWideningLayer, Constants.GreenBufferZoneLayer };

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


        Dictionary<ObjectId, string> roadlineDict = new Dictionary<ObjectId, string>();
        Dictionary<ObjectId, string> plotlineDict = new Dictionary<ObjectId, string>();

        // Get all LWPOLYLINE entities on _SurveyNo layer
        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        {
            BlockTable acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

            TotalSiteArea = GetAreaByLayer(Constants.PlotLayer, acTrans);

            PlotsArea = GetAreaByLayer(Constants.IndivPlotLayer, acTrans);
            //MortgageArea = GetAreaByLayer(Constants.MortgageLayer, acTrans);
            AmenitiesArea = GetAreaByLayer(Constants.AmenityLayer, acTrans);
            OpenSpaceArea = GetAreaByLayer(Constants.OpenSpaceLayer, acTrans);
            UtilityArea = GetAreaByLayer(Constants.UtilityLayer, acTrans);
            InternalRoadsArea = GetAreaByLayer(Constants.InternalRoadLayer, acTrans);
            SplayArea = GetAreaByLayer(Constants.SplayLayer, acTrans);
            LeftOverOwnerLandArea = GetAreaByLayer(Constants.LeftOverOwnerLandLayer, acTrans);
            RoadWideningArea = GetAreaByLayer(Constants.RoadWideningLayer, acTrans);
            GreenArea = GetAreaByLayer(Constants.GreenBufferZoneLayer, acTrans);

            VerifiedArea = PlotsArea + AmenitiesArea + OpenSpaceArea + UtilityArea + InternalRoadsArea + SplayArea + LeftOverOwnerLandArea + RoadWideningArea + GreenArea;

            differenceArea = TotalSiteArea - VerifiedArea;

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

            //ToDo

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

                                            if (uniquePoints.Count > uniquePointsIdentifier)
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
                                                    plotNo = FillPlotObject(surveyNos, acPoly2, plotNo, surveyNo, acTrans, Constants.IndivPlotLayer, Constants.IndivPlotDimLayer);

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

                                                surveyNo._PlotNos.Add(plotNo); //add plotNo to SurveyNo List
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
                                                plotNo = FillPlotObject(surveyNos, acPoly2, plotNo, surveyNo, acTrans, Constants.IndivPlotLayer, Constants.IndivPlotDimLayer);
                                                plotNo.AreaInSurveyNo.Add(surveyNo, plotNo._Area);
                                            }
                                            surveyNo._PlotNos.Add(plotNo); //add plotNo to SurveyNo List
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

                                            if (uniquePoints.Count > uniquePointsIdentifier)
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
                                                    amenityPlotNo = FillPlotObject(surveyNos, acPoly2, amenityPlotNo, surveyNo, acTrans, Constants.AmenityLayer, Constants.AmenityDimLayer);
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
                                                surveyNo._AmenityPlots.Add(amenityPlotNo); //add amenityPlot to SurveyNo List
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
                                                amenityPlotNo = FillPlotObject(surveyNos, acPoly2, amenityPlotNo, surveyNo, acTrans, Constants.AmenityLayer, Constants.AmenityDimLayer);
                                                amenityPlotNo.IsAmenity = true;
                                                amenityPlotNo.AreaInSurveyNo.Add(surveyNo, amenityPlotNo._Area);

                                            }
                                            surveyNo._AmenityPlots.Add(amenityPlotNo); //add amenityPlot to SurveyNo List
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

                List<Point3d> eastPointsCollection = new List<Point3d>();
                //Point3d point1 = new Point3d(item.eastPoints[0].X + 0.5, item.eastPoints[0].Y - 2, 0);
                //Point3d point2 = new Point3d(item.eastPoints[1].X + 0.5, item.eastPoints[1].Y - 2, 0);
                //Point3d point3 = new Point3d(item.eastPoints[0].X, item.eastPoints[0].Y - 2, 0);
                //Point3d point4 = new Point3d(item.eastPoints[1].X, item.eastPoints[1].Y - 2, 0);

                Point3d epoint1 = new Point3d(item.eastLineSegment[0].MidPoint.X + 0.5, item.eastLineSegment[0].MidPoint.Y + 0.5, 0);
                Point3d epoint2 = new Point3d(item.eastLineSegment[0].MidPoint.X - 0.5, item.eastLineSegment[0].MidPoint.Y - 0.5, 0);
                eastPointsCollection.AddRange(new List<Point3d> { epoint1, epoint2/*, point3, point4*/ });

                List<Polyline> roadPolylinesInEast = GetPolylinesUsingCrossPolygon(eastPointsCollection, acTrans, Constants.InternalRoadLayer);
                List<Polyline> plotPolylinesInEast = GetPolylinesUsingCrossPolygon(eastPointsCollection, acTrans, Constants.IndivPlotLayer);
                List<Polyline> amenityPolylinesInEast = GetPolylinesUsingCrossPolygon(eastPointsCollection, acTrans, Constants.AmenityLayer);

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
                }

                catch (System.Exception ee)
                {
                    //key not available in dictionary
                }

                #endregion

                #region Fill South Side Info
                //ed.WriteMessage("Filling South information...");

                List<Point3d> southPointsCollection = new List<Point3d>();
                Point3d spoint1 = new Point3d(item.southLineSegment[0].MidPoint.X + 0.5, item.southLineSegment[0].MidPoint.Y + 0.5, 0);
                Point3d spoint2 = new Point3d(item.southLineSegment[0].MidPoint.X - 0.5, item.southLineSegment[0].MidPoint.Y - 0.5, 0);
                southPointsCollection.AddRange(new List<Point3d> { spoint1, spoint2 });

                List<Polyline> roadPolylinesInSouth = GetPolylinesUsingCrossPolygon(southPointsCollection, acTrans, Constants.InternalRoadLayer);
                List<Polyline> plotPolylinesInSouth = GetPolylinesUsingCrossPolygon(southPointsCollection, acTrans, Constants.IndivPlotLayer);
                List<Polyline> amenityPolylinesInSouth = GetPolylinesUsingCrossPolygon(southPointsCollection, acTrans, Constants.AmenityLayer);

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
                }

                catch (System.Exception ee)
                {
                    //key not available in dictionary
                }

                #endregion

                #region Fill West Side Info
                //ed.WriteMessage("Filling West information...");

                List<Point3d> westPointsCollection = new List<Point3d>();
                Point3d wpoint1 = new Point3d(item.westLineSegment[0].MidPoint.X + 0.5, item.westLineSegment[0].MidPoint.Y + 0.5, 0);
                Point3d wpoint2 = new Point3d(item.westLineSegment[0].MidPoint.X - 0.5, item.westLineSegment[0].MidPoint.Y - 0.5, 0);
                westPointsCollection.AddRange(new List<Point3d> { wpoint1, wpoint2 });

                List<Polyline> roadPolylinesInWest = GetPolylinesUsingCrossPolygon(westPointsCollection, acTrans, Constants.InternalRoadLayer);
                List<Polyline> plotPolylinesInWest = GetPolylinesUsingCrossPolygon(westPointsCollection, acTrans, Constants.IndivPlotLayer);
                List<Polyline> amenityPolylinesInWest = GetPolylinesUsingCrossPolygon(westPointsCollection, acTrans, Constants.AmenityLayer);

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
                }

                catch (System.Exception ee)
                {
                    //key not available in dictionary
                }

                #endregion

                #region Fill North Side Info
                //ed.WriteMessage("Filling North information...");

                List<Point3d> northPointsCollection = new List<Point3d>();
                Point3d npoint1 = new Point3d(item.northLineSegment[0].MidPoint.X + 0.5, item.northLineSegment[0].MidPoint.Y + 0.5, 0);
                Point3d npoint2 = new Point3d(item.northLineSegment[0].MidPoint.X - 0.5, item.northLineSegment[0].MidPoint.Y - 0.5, 0);
                northPointsCollection.AddRange(new List<Point3d> { npoint1, npoint2 });

                List<Polyline> roadPolylinesInNorth = GetPolylinesUsingCrossPolygon(northPointsCollection, acTrans, Constants.InternalRoadLayer);
                List<Polyline> plotPolylinesInNorth = GetPolylinesUsingCrossPolygon(northPointsCollection, acTrans, Constants.IndivPlotLayer);
                List<Polyline> amenityPolylinesInNorth = GetPolylinesUsingCrossPolygon(northPointsCollection, acTrans, Constants.AmenityLayer);

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
                }

                catch (System.Exception ee)
                {
                    //key not available in dictionary
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

            // Write data to CSV

            UpdateAutoCADProgressBar(pm);

            DateTime datetime = DateTime.Now;
            string uniqueId = String.Format("{0:00}{1:00}{2:0000}{3:00}{4:00}{5:00}{6:000}",
                datetime.Day, datetime.Month, datetime.Year,
                datetime.Hour, datetime.Minute, datetime.Second, datetime.Millisecond);

            //System.Windows.Forms.MessageBox.Show(uniqueId);

            string fullfileName = Path.Combine(Path.GetDirectoryName(acCurDb.Filename), Path.GetFileNameWithoutExtension(acCurDb.Filename) + $"{"_" + uniqueId}");

            string csvFileNew = Path.Combine(Path.GetDirectoryName(acCurDb.Filename), Path.GetFileNameWithoutExtension(acCurDb.Filename) + $"{"_" + uniqueId}") + ".csv";

            string prefix = Path.GetFileNameWithoutExtension(acCurDb.Filename) + "_";
            string folderPath = Path.GetDirectoryName(acCurDb.Filename);

            WritetoCSV(csvFileNew, combinedPlots);

            ed.WriteMessage("Generating Report...");

            WritetoExcel(prefix, folderPath, combinedPlots);

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
        for (int i = 0; i < 10; i++)
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
                    if (item.AreaInSurveyNo[svno] > minArea)
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

            sw.WriteLine($",,,,,Total Site Area : " + TotalSiteArea);

        }
    }

    private void WritetoExcel(string prefix, string path, List<Plot> combinedPlots)
    {
        //VctDataTableRepository repo = new VctDataTableRepository();
        //repo.TemplatePath = @"C:\Data\Square_Excel_Template.xlsx";
        //repo.Prefix = prefix;
        //repo.SavePath = path;
        //if (!Directory.Exists(repo.SavePath))
        //    Directory.CreateDirectory(repo.SavePath);
        //repo.MergePdf = true;
        //repo.OpenExcelAfterGenerate = true;
        //repo.OpenPdfAfterGenerate = true;

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
                if (item.AreaInSurveyNo[svno] > minArea)
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

        dt1.Rows.Add(new object[] {$"" ,
               $"" ,
               $"" ,
               $"" ,
               $"" ,
               $"Total Site Area : {TotalSiteArea}" });

        WritetoExcelAndPDF.WritetoExcel2(prefix, path, dt1);


        //dt1.Rows.Add(new object[] { "James Bond, LLC", 120, "Garrison" });
        //dt1.Rows.Add(new object[] { "LLC", 10, "Gar" });
        //dt1.Rows.Add(new object[] { "Bond, LLC", 10, "Gar" });

        //var highlighter = new StyleSettings()
        //{
        //    BackColor = new HighlightColor() { B = 144, G = 238, R = 144 }
        //};

        //VctDataTable dataTable1 = new VctDataTable(dt1)
        //{
        //    SheetName = "Meters",
        //    PrintHeader = false,
        //    StartRow = 3
        //};

        ////dataTable1.Rows[0].Cells[0].UpdateSettings = true;
        ////dataTable1.Rows[0].Cells[0].Settings = highlighter;
        ////dataTable1.Rows[2].Cells[2].UpdateSettings = true;
        ////dataTable1.Rows[2].Cells[2].Settings = highlighter;

        //repo.VctDataTables.Add(dataTable1);
        //repo.GenerateReport(true);
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

    private SelectionFilter CreateSelectionFilterByStartTypeAndLayer(string startType1, string startType2, string Layer)
    {
        TypedValue[] Filter = { new TypedValue((int)DxfCode.Start, startType1),
                                new TypedValue((int)DxfCode.Start, startType2),
                                new TypedValue((int)DxfCode.LayerName, Layer)
                              };

        SelectionFilter acSelFtr = new SelectionFilter(Filter);
        return acSelFtr;
    }

    private Plot FillPlotObject(List<SurveyNo> surveyNos, Polyline acPoly2, Plot plotNo, SurveyNo surveyNo, Transaction acTrans, string TextLayerName, string dimensionLayerName)
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
            plotNo._Area = Math.Round(plotNo._Polyline.Area, AreaDecimals);
            plotNo._ParentSurveyNos.Add(surveyNo);
        }
        else
        {
            plotNo._PlotNo = GetTextFromLayer(acTrans, plotNo._PolylinePoints, Constants.MTEXT, TextLayerName);//assign plotNo Text
            plotNo._Area = Math.Round(plotNo._Polyline.Area, AreaDecimals);
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

        //get all dimensions of plot and fill in plot object
        PromptSelectionResult dimSelResult = ed.SelectCrossingPolygon(plotNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer("DIMENSION", dimensionLayerName));

        if (dimSelResult.Status == PromptStatus.OK)
        {
            SelectionSet dimSelSet = dimSelResult.Value;
            foreach (SelectedObject acSSObj in dimSelSet)
            {
                if (acSSObj != null)
                {
                    Dimension dimEntity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Dimension;
                    if (dimEntity != null)
                    {
                        string dimValue = dimEntity.DimensionText;

                        if (Convert.ToDouble(dimValue) > 0.5)
                            plotNo._AllDims.Add(new SDimension(dimEntity, dimEntity.DimensionText, dimEntity.TextPosition));
                    }
                }
            }
        }

        //fill sizes by direction
        FillLineSegmentsAndPointsByDirection(plotNo._Polyline, plotNo);
        FillSizesByDirection(plotNo);
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
        //filling sizes by direction based on available dimensions
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

    private void FillSizesByDirection(Plot plotNo, Transaction acTrans)
    {
        //plotNo._SizeInEast = GetSize(plotNo, acTrans, new Point3dCollection(plotNo.eastPoints.ToArray()));
        //plotNo._SizeInSouth = GetSize(plotNo, acTrans, new Point3dCollection(plotNo.southPoints.ToArray()));
        //plotNo._SizeInWest = GetSize(plotNo, acTrans, new Point3dCollection(plotNo.westPoints.ToArray()));
        //plotNo._SizeInNorth = GetSize(plotNo, acTrans, new Point3dCollection(plotNo.northPoints.ToArray()));
    }

    private string GetSize(Plot plotNo, Transaction acTrans, Point3dCollection point3dCollection)
    {
        PromptSelectionResult textSelResult = ed.SelectCrossingPolygon(point3dCollection, CreateSelectionFilterByStartTypeAndLayer("DIMENSION", Constants.IndivPlotDimLayer));

        if (textSelResult.Status == PromptStatus.OK)
        {
            Dimension textEntity = acTrans.GetObject(textSelResult.Value[0].ObjectId, OpenMode.ForRead) as Dimension;
            if (textEntity != null)
            {
                string fval2 = textEntity.DimensionText;
                return fval2;
            }
        }

        return "";
    }

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
