using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;


namespace AcadProject
{
    public class Commands
    {
        [CommandMethod("HW")]
        public void HelloWorld()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            ed.WriteMessage("Hello, AutoCAD!");
            System.Windows.Forms.MessageBox.Show("Hello, AutoCAD!", "AutoCAD Message");
        }

        [CommandMethod("CreateLine")]
        public void CreateLine()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                Line line = new Line(new Point3d(0, 0, 0), new Point3d(10, 10, 0));
                btr.AppendEntity(line);
                trans.AddNewlyCreatedDBObject(line, true);

                trans.Commit();
            }
        }

        [CommandMethod("GetLWPOLYLINEEntities")]
        public void GetLWPOLYLINEEntities()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for read
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                // Define a TypedValue array to filter for LWPOLYLINE entities on the _SurveyNo layer
                TypedValue[] acTypValAr = new TypedValue[2];
                acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 0);
                acTypValAr.SetValue(new TypedValue((int)DxfCode.LayerName, "_SurveyNo"), 1);

                // Assign the filter criteria to a SelectionFilter object
                SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);

                // Request for objects to be selected in the drawing area
                PromptSelectionResult acSSPrompt = acDoc.Editor.SelectAll(acSelFtr);

                // If the prompt status is OK, objects were selected
                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt.Value;

                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        if (acSSObj != null)
                        {
                            Entity acEnt = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead) as Entity;

                            if (acEnt is Polyline)
                            {
                                Polyline acPoly = acEnt as Polyline;
                                // Process the LWPOLYLINE entity
                                acDoc.Editor.WriteMessage("\nLWPOLYLINE found with ObjectId: " + acPoly.ObjectId);
                            }
                        }
                    }
                }

                // Dispose of the transaction
                acTrans.Commit();
            }
        }
    }
}
