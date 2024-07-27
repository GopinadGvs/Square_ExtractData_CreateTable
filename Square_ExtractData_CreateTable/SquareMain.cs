using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

[assembly: CommandClass(typeof(MyCommands))]

public class MyCommands
{
    [CommandMethod("SIP")]
    public void SIP()
    {
        Document acDoc = Application.DocumentManager.MdiActiveDocument;
        Database acCurDb = acDoc.Database;
        Editor ed = acDoc.Editor;

        // Zoom extents
        ed.Command("_zoom", "_e");

        // Turn on and thaw layers
        ed.Command("_-layer", "t", "_SurveyNo", "", "ON", "_SurveyNo", "", "t", "_IndivSubPlot", "", "ON", "_IndivSubPlot", "");

        List<(string, ObjectId)> snoPno = new List<(string, ObjectId)>();
        List<(string, string)> snoPnoVal = new List<(string, string)>();

        // Get all LWPOLYLINE entities on _SurveyNo layer
        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        {
            BlockTable acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

            TypedValue[] acTypValAr = new TypedValue[]
            {
                new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                new TypedValue((int)DxfCode.LayerName, "_SurveyNo")
            };
            SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
            PromptSelectionResult acSSPrompt = ed.SelectAll(acSelFtr);

            if (acSSPrompt.Status == PromptStatus.OK)
            {
                SelectionSet acSSet = acSSPrompt.Value;

                foreach (SelectedObject acSSObj in acSSet)
                {
                    if (acSSObj != null)
                    {
                        Polyline acPoly = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Polyline;

                        if (acPoly != null)
                        {
                            Point3dCollection pts = new Point3dCollection();
                            for (int i = 0; i < acPoly.NumberOfVertices; i++)
                            {
                                pts.Add(acPoly.GetPoint3dAt(i));
                            }

                            // Find text entity on _SurveyNo layer
                            TypedValue[] acTypValArText = new TypedValue[]
                            {
                                new TypedValue((int)DxfCode.Start, "TEXT"),
                                new TypedValue((int)DxfCode.LayerName, "_SurveyNo")
                            };
                            SelectionFilter acSelFtrText = new SelectionFilter(acTypValArText);
                            PromptSelectionResult acSSPromptText = ed.SelectCrossingPolygon(pts, acSelFtrText);

                            if (acSSPromptText.Status == PromptStatus.OK)
                            {
                                SelectionSet acSSetText = acSSPromptText.Value;
                                DBText acText = acTrans.GetObject(acSSetText[0].ObjectId, OpenMode.ForRead) as DBText;
                                string val2 = acText.TextString;

                                // Find intersecting polylines on the _IndivSubPlot layer
                                TypedValue[] acTypValArPoly = new TypedValue[]
                                {
                                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                                    new TypedValue((int)DxfCode.LayerName, "_IndivSubPlot")
                                };
                                SelectionFilter acSelFtrPoly = new SelectionFilter(acTypValArPoly);
                                PromptSelectionResult acSSPromptPoly = ed.SelectCrossingPolygon(pts, acSelFtrPoly);

                                if (acSSPromptPoly.Status == PromptStatus.OK)
                                {
                                    SelectionSet acSSetPoly = acSSPromptPoly.Value;
                                    foreach (SelectedObject acSSObjPoly in acSSetPoly)
                                    {
                                        if (acSSObjPoly != null)
                                        {
                                            Polyline acPoly2 = acTrans.GetObject(acSSObjPoly.ObjectId, OpenMode.ForRead) as Polyline;
                                            if (acPoly2 != null)
                                            {
                                                if (acPoly2.GeometricExtents != null)
                                                {
                                                    snoPno.Add((val2, acPoly2.ObjectId));
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            acTrans.Commit();
        }

        // Get unique snoPno values and sort
        var snoPnoUnique = snoPno.Distinct().OrderBy(x => x.Item1).ToList();

        foreach (var val in snoPnoUnique)
        {
            string val2 = "";
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                DBText acText = acTrans.GetObject(val.Item2, OpenMode.ForRead) as DBText;
                if (acText != null)
                {
                    val2 = acText.TextString;
                    snoPnoVal.Add((val2, val.Item1));
                }

                acTrans.Commit();
            }
        }

        snoPnoVal = snoPnoVal.OrderBy(x => int.Parse(x.Item1)).ToList();

        // Write data to CSV
        string csvFileNew = Path.Combine(acCurDb.Filename, "SIP.csv");

        using (StreamWriter sw = new StreamWriter(csvFileNew))
        {
            sw.WriteLine("Plot Number,Survey No");
            foreach (var itm in snoPnoVal)
            {
                sw.WriteLine($"{itm.Item1},{itm.Item2}");
            }
        }

        System.Diagnostics.Process.Start("notepad.exe", csvFileNew);

        // Turn off _SurveyNo layer
        ed.Command("_-layer", "OFF", "_SurveyNo", "");

        ed.WriteMessage("\nProcess complete.");
    }
}
