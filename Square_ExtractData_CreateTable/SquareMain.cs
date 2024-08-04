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

[assembly: CommandClass(typeof(MyCommands))]

public class MyCommands
{
    private static Editor ed = null;
    List<string> mortgagePlots = new List<string>();
    private static int AreaDecimals = 2;

    [CommandMethod("MSIP")]
    public void SIP()
    {
        Document acDoc = Application.DocumentManager.MdiActiveDocument;

        LogWriter.LogWrite(acDoc.Name);

        Database acCurDb = acDoc.Database;
        /*Editor*/
        ed = acDoc.Editor;

        // Zoom extents
        ed.Command("_.zoom", "_e");

        // Turn on and thaw layers
        //ed.Command("_-layer", "t", "_SurveyNo", "", "ON", "_SurveyNo", "", "t", "_IndivSubPlot", "", "ON", "_IndivSubPlot", "");

        //List<(string, ObjectId)> snoPno = new List<(string, ObjectId)>(); -> old
        //List<(string, string)> snoPnoVal = new List<(string, string)>(); -> old

        List<SurveyNo> surveyNos = new List<SurveyNo>();
        List<Mortgage> mortgages = new List<Mortgage>();

        // Get all LWPOLYLINE entities on _SurveyNo layer
        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        {
            BlockTable acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

            #region Logic to get mortgage plot numbers list

            PromptSelectionResult acSSPrompt1 = ed.SelectAll(CreateSelectionFilterByStartTypeAndLayer("LWPOLYLINE", "_MortgageArea"));

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
                            List<string> plotsInsideMortgage = GetListTextFromLayer(acTrans, mortgage._PolylinePoints, "TEXT", "_IndivSubPlot");
                            if (plotsInsideMortgage.Count > 0)
                            {
                                mortgage._PlotNos.AddRange(plotsInsideMortgage);
                            }
                            else
                            {
                                mortgage._PlotNos.AddRange(GetListTextFromLayer(acTrans, mortgage._PolylinePoints, "MTEXT", "_IndivSubPlot"));
                            }

                            #region Old code
                            //PromptSelectionResult acSSPromptText = ed.SelectCrossingPolygon(mortgage._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer("TEXT", "_IndivSubPlot"));

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

            #region Process SurveyNo Polylines

            PromptSelectionResult acSSPrompt = ed.SelectAll(CreateSelectionFilterByStartTypeAndLayer("LWPOLYLINE", "_SurveyNo"));

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
                            string surveyNoText = GetTextFromLayer(acTrans, surveyNo._PolylinePoints, "TEXT", "_SurveyNo");
                            if (!string.IsNullOrEmpty(surveyNoText))
                            {
                                surveyNo._SurveyNo = surveyNoText;//assign surveyNo
                            }
                            else
                            {
                                surveyNo._SurveyNo = GetTextFromLayer(acTrans, surveyNo._PolylinePoints, "MTEXT", "_SurveyNo");
                            }

                            //Get DocNo Text info
                            string DocNoText = GetTextFromLayer(acTrans, surveyNo._PolylinePoints, "TEXT", "_DocNo");
                            if (!string.IsNullOrEmpty(DocNoText))
                            {
                                surveyNo.DocumentNo = DocNoText;//assign DocNo
                            }
                            else
                            {
                                surveyNo.DocumentNo = GetTextFromLayer(acTrans, surveyNo._PolylinePoints, "MTEXT", "_DocNo");
                            }

                            //Get LandLord Text info
                            string LandLordText = GetTextFromLayer(acTrans, surveyNo._PolylinePoints, "TEXT", "_LandLord");
                            if (!string.IsNullOrEmpty(LandLordText))
                            {
                                surveyNo.LandLordName = LandLordText;//assign LandLordName
                            }
                            else
                            {
                                surveyNo.LandLordName = GetTextFromLayer(acTrans, surveyNo._PolylinePoints, "MTEXT", "_LandLord");
                            }

                            #endregion

                            #region Collect & Fill IndivSubPlot data using Cross Polygon                                

                            PromptSelectionResult acSSPromptPoly = ed.SelectCrossingPolygon(surveyNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer("LWPOLYLINE", "_IndivSubPlot"));

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

                                            if (uniquePoints.Count > 1)
                                            {
                                                // new logic added
                                                var existingPlotNos = surveyNos.SelectMany(x => x._PlotNos).Where
                                                (x => x._Polyline.ObjectId == acPoly2.ObjectId).ToList();

                                                if (existingPlotNos.Count > 0)
                                                {
                                                    plotNo = existingPlotNos[0];
                                                    plotNo._ParentSurveyNos.Add(surveyNo);
                                                }

                                                else
                                                {
                                                    //considering dimensions from layer "_IndivSubPlot_DIMENSION"
                                                    plotNo = FillPlotObject(surveyNos, acPoly2, plotNo, surveyNo, acTrans, "_IndivSubPlot", "_IndivSubPlot_DIMENSION");

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
                            PromptSelectionResult acSSPromptZeroPoly = ed.SelectWindowPolygon(surveyNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer("LWPOLYLINE", "_IndivSubPlot"));

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
                                            }

                                            else
                                            {
                                                //considering dimensions from layer "_IndivSubPlot_DIMENSION"
                                                plotNo = FillPlotObject(surveyNos, acPoly2, plotNo, surveyNo, acTrans, "_IndivSubPlot", "_IndivSubPlot_DIMENSION");

                                            }
                                            surveyNo._PlotNos.Add(plotNo); //add plotNo to SurveyNo List
                                        }
                                    }
                                }
                            }

                            #endregion

                            #region Collect & Fill Amenity data using Cross Polygon                                

                            PromptSelectionResult acSSPromptPoly1 = ed.SelectCrossingPolygon(surveyNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer("LWPOLYLINE", "_Amenity"));

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

                                            if (uniquePoints.Count > 1)
                                            {
                                                var existingPlotNos = surveyNos.SelectMany(x => x._AmenityPlots).Where
                                            (x => x._Polyline.ObjectId == acPoly2.ObjectId).ToList();

                                                if (existingPlotNos.Count > 0)
                                                {
                                                    amenityPlotNo = existingPlotNos[0];
                                                    amenityPlotNo._ParentSurveyNos.Add(surveyNo);
                                                }

                                                else
                                                {
                                                    //considering dimensions from layer "_Amenity_DIMENSION"
                                                    amenityPlotNo = FillPlotObject(surveyNos, acPoly2, amenityPlotNo, surveyNo, acTrans, "_Amenity", "_Amenity_DIMENSION");
                                                    amenityPlotNo.IsAmenity = true;

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
                            PromptSelectionResult acSSPromptZeroPoly2 = ed.SelectWindowPolygon(surveyNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer("LWPOLYLINE", "_Amenity"));

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
                                            }

                                            else
                                            {
                                                //considering dimensions from layer "_Amenity_DIMENSION"                                                    
                                                amenityPlotNo = FillPlotObject(surveyNos, acPoly2, amenityPlotNo, surveyNo, acTrans, "_Amenity", "_Amenity_DIMENSION");
                                                amenityPlotNo.IsAmenity = true;

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

            acTrans.Commit();
        }

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
        //                new TypedValue((int)DxfCode.Start, "TEXT"),
        //                new TypedValue((int)DxfCode.LayerName, "_IndivSubPlot")
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
        //        new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
        //        new TypedValue((int)DxfCode.LayerName, "_IndivSubPlot")
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
        //                        new TypedValue((int)DxfCode.Start, "TEXT"),
        //                        new TypedValue((int)DxfCode.LayerName, "_IndivSubPlot")
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

        DateTime datetime = DateTime.Now;
        string uniqueId = String.Format("{0:00}{1:00}{2:0000}{3:00}{4:00}{5:00}{6:000}",
            datetime.Day, datetime.Month, datetime.Year,
            datetime.Hour, datetime.Minute, datetime.Second, datetime.Millisecond);

        //System.Windows.Forms.MessageBox.Show(uniqueId);


        string csvFileNew = Path.Combine(Path.GetDirectoryName(acCurDb.Filename), Path.GetFileNameWithoutExtension(acCurDb.Filename) + $"{"_" + uniqueId + ".csv" }"
        );

        using (StreamWriter sw = new StreamWriter(csvFileNew))
        {
            sw.WriteLine("Plot Number,East,South,West,North,Plot Area, Mortgage Plots, Amenity Plots,Doc.No/R.S.No./Area/Name");

            foreach (var item in combinedPlots)
            {
                sw.WriteLine($"{item._PlotNo}," +
                    $"{item._SizesInEast[0].Text}," +
                    $"{item._SizesInSouth[0].Text}," +
                    $"{item._SizesInWest[0].Text}," +
                    $"{item._SizesInNorth[0].Text}," +
                    //$"{Convert.ToString(string.Join("|", item._ParentSurveyNos.Select(x => x._SurveyNo).ToArray()))}," +
                    $"{(/*item._PlotArea == 0 ? "" : */item._PlotArea.ToString())}," +
                    $"{(/*item._MortgageArea == 0 ? "" : */item._MortgageArea.ToString())}," +
                    $"{(/*item._AmenityArea == 0 ? "" : */item._AmenityArea.ToString())}"
                    );
            }

            sw.WriteLine($"," +
                   $"," +
                   $"," +
                   $"," +
                   $"," +
                   //$"," +
                   $"{combinedPlots.Select(x => x._PlotArea).ToArray().Sum().ToString()}," +
                   $"{combinedPlots.Select(x => x._MortgageArea).ToArray().Sum().ToString()}," +
                   $"{combinedPlots.Select(x => x._AmenityArea).ToArray().Sum().ToString()}"
                   );
        }


        #region Test write

        //sw.WriteLine("Plot Number,East,South,West,North,Survey No,Center,EP1,EP2,SP1,SP2,WP1,WP2,NP1,NP2");

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
        System.Diagnostics.Process.Start("Excel.exe", csvFileNew);

        // Turn off _SurveyNo layer
        //ed.Command("_-layer", "OFF", "_SurveyNo", "");

        ed.WriteMessage("\nProcess complete.");
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
                {
                    myObject.northLineSegment.Add(polyline.GetLineSegmentAt(i));
                    myObject.northPoints.Add(vertex1);
                    myObject.northPoints.Add(vertex2);
                }

                if ((direction1.X < 0 && direction1.Y < 0) && (direction2.X > 0 && direction2.Y < 0) ||
                    (direction2.X < 0 && direction2.Y < 0) && (direction1.X > 0 && direction1.Y < 0))
                {
                    myObject.southLineSegment.Add(polyline.GetLineSegmentAt(i));
                    myObject.southPoints.Add(vertex1);
                    myObject.southPoints.Add(vertex2);
                }

                if ((direction1.X > 0 && direction1.Y > 0) && (direction2.X > 0 && direction2.Y < 0) ||
                    (direction2.X > 0 && direction2.Y > 0) && (direction1.X > 0 && direction1.Y < 0))
                {
                    myObject.eastLineSegment.Add(polyline.GetLineSegmentAt(i));
                    myObject.eastPoints.Add(vertex1);
                    myObject.eastPoints.Add(vertex2);
                }

                if ((direction1.X < 0 && direction1.Y > 0) && (direction2.X < 0 && direction2.Y < 0) ||
                    (direction2.X < 0 && direction2.Y > 0) && (direction1.X < 0 && direction1.Y < 0))
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
        string plotNoText = GetTextFromLayer(acTrans, plotNo._PolylinePoints, "TEXT", TextLayerName);
        if (!string.IsNullOrEmpty(plotNoText))
        {
            plotNo._PlotNo = plotNoText;//assign plotNo Text
            plotNo._Area = Math.Round(plotNo._Polyline.Area, AreaDecimals);
            plotNo._ParentSurveyNos.Add(surveyNo);
        }
        else
        {
            plotNo._PlotNo = GetTextFromLayer(acTrans, plotNo._PolylinePoints, "MTEXT", TextLayerName);//assign plotNo Text
            plotNo._Area = Math.Round(plotNo._Polyline.Area, AreaDecimals);
            plotNo._ParentSurveyNos.Add(surveyNo);
        }

        #region Old Code
        //fill plot number text assuming text is of single text, area and parent survey No
        //PromptSelectionResult textSelResult = ed.SelectWindowPolygon(plotNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer("TEXT", TextLayerName));

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
        //PromptSelectionResult textSelResult1 = ed.SelectWindowPolygon(plotNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer("MTEXT", TextLayerName));

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
        PromptSelectionResult textSelResult = ed.SelectCrossingPolygon(point3dCollection, CreateSelectionFilterByStartTypeAndLayer("DIMENSION", "_IndivSubPlot_DIMENSION"));

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

//PromptSelectionResult mortgageSelResult = ed.SelectCrossingPolygon(plotNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer("LWPOLYLINE", "_MortgageArea"));

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
