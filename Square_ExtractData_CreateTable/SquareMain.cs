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


    [CommandMethod("MSIP")]
    public void SIP()
    {
        Document acDoc = Application.DocumentManager.MdiActiveDocument;
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

        // Get all LWPOLYLINE entities on _SurveyNo layer
        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        {
            BlockTable acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

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
                            FillAllPointsAndByDirection(surveyNo._Polyline, surveyNo);

                            PromptSelectionResult acSSPromptText = ed.SelectCrossingPolygon(surveyNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer("TEXT", "_SurveyNo"));

                            if (acSSPromptText.Status == PromptStatus.OK)
                            {
                                SelectionSet acSSetText = acSSPromptText.Value;
                                DBText acText = acTrans.GetObject(acSSetText[0].ObjectId, OpenMode.ForRead) as DBText;
                                string val2 = acText.TextString;

                                surveyNo._SurveyNo = val2; //assign surveyNo

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
                                                    FillPlotObject(surveyNos, acPoly2, plotNo, surveyNo/*, ed*/, acTrans);
                                                }
                                            }
                                        }
                                    }
                                }

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
                                                FillPlotObject(surveyNos, acPoly2, plotNo, surveyNo/*, ed*/, acTrans);
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        surveyNos.Add(surveyNo); //add surveyNo to list
                    }
                }
            }

            acTrans.Commit();
        }

        // Select, Distinct, and Sort (using comparer to sort strings like 1,10,2,20,20A,3,30,30C,4)
        var uniquePlots = surveyNos
            .SelectMany(x => x._PlotNos)
            .Distinct()
            .OrderBy(x => x, new AlphanumericPlotComparer())
            .ToList();



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
        string csvFileNew = Path.Combine(Path.GetDirectoryName(acCurDb.Filename), Path.GetFileNameWithoutExtension(acCurDb.Filename) + ".csv");

        using (StreamWriter sw = new StreamWriter(csvFileNew))
        {
            sw.WriteLine("Plot Number,East,South,West,North,Survey No");
            //sw.WriteLine("Plot Number,East,South,West,North,Survey No,Center,EP1,EP2,SP1,SP2,WP1,WP2,NP1,NP2");

            foreach (var item in uniquePlots)
            {
                sw.WriteLine($"{item._PlotNo}," +
                    $"{item._SizesInEast[0].Text}," +
                    $"{item._SizesInSouth[0].Text}," +
                    $"{item._SizesInWest[0].Text}," +
                    $"{item._SizesInNorth[0].Text}," +
                    $"{Convert.ToString(string.Join("|", item._ParentSurveyNos.Select(x => x._SurveyNo).ToArray()))}"
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
                    );
            }
        }

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

    private void FillAllPointsAndByDirection(Polyline polyline, dynamic myObject)
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

        // Iterate through the polyline vertices
        for (int i = 0; i < polyline.NumberOfVertices; i++)
        {
            myObject._PolylinePoints.Add(polyline.GetPoint3dAt(i));

            Point3d vertex = polyline.GetPoint3dAt(i);

            Vector3d direction = vertex - myObject.Center;

            if (direction.Y > 0)
            {
                myObject.northPoints.Add(vertex);
            }
            else if (direction.Y < 0)
            {
                myObject.southPoints.Add(vertex);
            }

            if (direction.X > 0)
            {
                myObject.eastPoints.Add(vertex);
            }
            else if (direction.X < 0)
            {
                myObject.westPoints.Add(vertex);
            }
        }


        //ToDo -> Issue with Circular curves

        // Convert the dynamic lists to List<Point3d> and then process them
        List<Point3d> northPoints = ((IEnumerable<Point3d>)myObject.northPoints).OrderBy(p => p.DistanceTo((Point3d)myObject.Center)).Take(2).ToList();
        List<Point3d> southPoints = ((IEnumerable<Point3d>)myObject.southPoints).OrderBy(p => p.DistanceTo((Point3d)myObject.Center)).Take(2).ToList();
        List<Point3d> eastPoints = ((IEnumerable<Point3d>)myObject.eastPoints).OrderBy(p => p.DistanceTo((Point3d)myObject.Center)).Take(2).ToList();
        List<Point3d> westPoints = ((IEnumerable<Point3d>)myObject.westPoints).OrderBy(p => p.DistanceTo((Point3d)myObject.Center)).Take(2).ToList();

        // Assign the processed lists back to the dynamic object
        myObject.northPoints = northPoints;
        myObject.southPoints = southPoints;
        myObject.eastPoints = eastPoints;
        myObject.westPoints = westPoints;

    }

    private SelectionFilter CreateSelectionFilterByStartTypeAndLayer(string startType, string Layer)
    {
        TypedValue[] Filter = { new TypedValue((int)DxfCode.Start, startType),
                                new TypedValue((int)DxfCode.LayerName, Layer)
                              };

        SelectionFilter acSelFtr = new SelectionFilter(Filter);
        return acSelFtr;
    }

    private void FillPlotObject(List<SurveyNo> surveyNos, Polyline acPoly2, Plot plotNo, SurveyNo surveyNo/*, Editor ed*/, Transaction acTrans)
    {
        // ToDo Add new logic here
        var existingPlotNos = surveyNos.SelectMany(x => x._PlotNos).Where
        (x => x._Polyline.ObjectId == acPoly2.ObjectId).ToList();

        if (existingPlotNos.Count > 0)
        {
            plotNo = existingPlotNos[0];
            plotNo._ParentSurveyNos.Add(surveyNo);
            //FillSizesByDirection(plotNo, acTrans);
        }

        else
        {
            plotNo._Polyline = acPoly2; //assign polyline
            plotNo.Center = getCenter(plotNo._Polyline);
            FillAllPointsAndByDirection(plotNo._Polyline, plotNo);

            PromptSelectionResult textSelResult = ed.SelectWindowPolygon(plotNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer("TEXT", "_IndivSubPlot"));

            if (textSelResult.Status == PromptStatus.OK)
            {
                DBText textEntity = acTrans.GetObject(textSelResult.Value[0].ObjectId, OpenMode.ForRead) as DBText;
                if (textEntity != null)
                {
                    string fval2 = textEntity.TextString;
                    plotNo._PlotNo = fval2;
                    plotNo._ParentSurveyNos.Add(surveyNo);
                    //FillSizesByDirection(plotNo, acTrans);                    
                }
            }

            PromptSelectionResult dimSelResult = ed.SelectCrossingPolygon(plotNo._PolylinePoints, CreateSelectionFilterByStartTypeAndLayer("DIMENSION", "_IndivSubPlot_DIMENSION"));

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

            FillSizesByDirection(plotNo);
        }

        surveyNo._PlotNos.Add(plotNo); //add plotNo to SurveyNo List
    }


    private void FillSizesByDirection(Plot plotNo)
    {
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

        //ToDo -> Issue with Circular curves

        plotNo._SizesInNorth = plotNo._SizesInNorth.OrderBy(p => p.position.DistanceTo(plotNo.Center)).Take(2).ToList();
        plotNo._SizesInSouth = plotNo._SizesInSouth.OrderBy(p => p.position.DistanceTo(plotNo.Center)).Take(2).ToList();
        plotNo._SizesInEast = plotNo._SizesInEast.OrderBy(p => p.position.DistanceTo(plotNo.Center)).Take(2).ToList();
        plotNo._SizesInWest = plotNo._SizesInWest.OrderBy(p => p.position.DistanceTo(plotNo.Center)).Take(2).ToList();

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
}
