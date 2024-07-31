using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Square_ExtractData_CreateTable
{
    class Snippets
    {

        #region calculate center of polyline

        //using Autodesk.AutoCAD.ApplicationServices;
        //using Autodesk.AutoCAD.DatabaseServices;
        //using Autodesk.AutoCAD.Geometry;
        //using Autodesk.AutoCAD.Runtime;

        //public class PolylineCenter
        //    {
        //        [CommandMethod("GetPolylineCenter")]
        //        public void GetPolylineCenter()
        //        {
        //            Document acDoc = Application.DocumentManager.MdiActiveDocument;
        //            Database acCurDb = acDoc.Database;

        //            // Start a transaction
        //            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        //            {
        //                // Open the Block table for read
        //                BlockTable acBlkTbl;
        //                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

        //                // Open the Block table record Model space for read
        //                BlockTableRecord acBlkTblRec;
        //                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

        //                // Create a PromptEntityOptions object to prompt the user to select a polyline
        //                PromptEntityOptions peo = new PromptEntityOptions("\nSelect a polyline: ");
        //                peo.SetRejectMessage("\nSelected entity is not a polyline.");
        //                peo.AddAllowedClass(typeof(Polyline), false);

        //                // Get the selected polyline
        //                PromptEntityResult per = acDoc.Editor.GetEntity(peo);
        //                if (per.Status != PromptStatus.OK)
        //                {
        //                    acTrans.Abort();
        //                    return;
        //                }

        //                // Open the selected polyline for read
        //                Polyline acPoly = acTrans.GetObject(per.ObjectId, OpenMode.ForRead) as Polyline;

        //                // Get the geometric extents of the polyline
        //                Extents3d polyExtents = acPoly.GeometricExtents;

        //                // Calculate the center point of the polyline
        //                Point3d centerPoint = new Point3d(
        //                    (polyExtents.MinPoint.X + polyExtents.MaxPoint.X) / 2,
        //                    (polyExtents.MinPoint.Y + polyExtents.MaxPoint.Y) / 2,
        //                    (polyExtents.MinPoint.Z + polyExtents.MaxPoint.Z) / 2);

        //                // Output the center point to the command line
        //                acDoc.Editor.WriteMessage($"\nCenter of the polyline: {centerPoint}");

        //                // Commit the transaction
        //                acTrans.Commit();
        //            }
        //        }
        //    }

        #endregion


        #region Logic to get segments in all directions

        //        using Autodesk.AutoCAD.ApplicationServices;
        //using Autodesk.AutoCAD.DatabaseServices;
        //using Autodesk.AutoCAD.Geometry;
        //using Autodesk.AutoCAD.Runtime;
        //using System.Collections.Generic;

        //public class PolylineDirections
        //    {
        //        [CommandMethod("GetPolylineSegmentsByDirection")]
        //        public void GetPolylineSegmentsByDirection()
        //        {
        //            Document acDoc = Application.DocumentManager.MdiActiveDocument;
        //            Database acCurDb = acDoc.Database;

        //            // Start a transaction
        //            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        //            {
        //                // Create a PromptEntityOptions object to prompt the user to select a polyline
        //                PromptEntityOptions peo = new PromptEntityOptions("\nSelect a polyline: ");
        //                peo.SetRejectMessage("\nSelected entity is not a polyline.");
        //                peo.AddAllowedClass(typeof(Polyline), false);

        //                // Get the selected polyline
        //                PromptEntityResult per = acDoc.Editor.GetEntity(peo);
        //                if (per.Status != PromptStatus.OK)
        //                {
        //                    acTrans.Abort();
        //                    return;
        //                }

        //                // Open the selected polyline for read
        //                Polyline acPoly = acTrans.GetObject(per.ObjectId, OpenMode.ForRead) as Polyline;

        //                // Define lists to hold segments in different directions
        //                List<LineSegment2d> northSegments = new List<LineSegment2d>();
        //                List<LineSegment2d> southSegments = new List<LineSegment2d>();
        //                List<LineSegment2d> eastSegments = new List<LineSegment2d>();
        //                List<LineSegment2d> westSegments = new List<LineSegment2d>();

        //                // Iterate through the polyline segments
        //                for (int i = 0; i < acPoly.NumberOfVertices; i++)
        //                {
        //                    if (acPoly.GetSegmentType(i) == SegmentType.Line)
        //                    {
        //                        LineSegment2d segment = acPoly.GetLineSegment2dAt(i);

        //                        // Determine the direction of the segment
        //                        if (segment.StartPoint.X == segment.EndPoint.X)
        //                        {
        //                            // Vertical line
        //                            if (segment.StartPoint.Y < segment.EndPoint.Y)
        //                            {
        //                                northSegments.Add(segment);
        //                            }
        //                            else
        //                            {
        //                                southSegments.Add(segment);
        //                            }
        //                        }
        //                        else if (segment.StartPoint.Y == segment.EndPoint.Y)
        //                        {
        //                            // Horizontal line
        //                            if (segment.StartPoint.X < segment.EndPoint.X)
        //                            {
        //                                eastSegments.Add(segment);
        //                            }
        //                            else
        //                            {
        //                                westSegments.Add(segment);
        //                            }
        //                        }
        //                    }
        //                }

        //                // Output the results
        //                acDoc.Editor.WriteMessage($"\nNorth Segments: {northSegments.Count}");
        //                acDoc.Editor.WriteMessage($"\nSouth Segments: {southSegments.Count}");
        //                acDoc.Editor.WriteMessage($"\nEast Segments: {eastSegments.Count}");
        //                acDoc.Editor.WriteMessage($"\nWest Segments: {westSegments.Count}");

        //                // Commit the transaction
        //                acTrans.Commit();
        //            }
        //        }
        //    }
        #endregion

        #region FillPointByDirection

        //        using Autodesk.AutoCAD.ApplicationServices;
        //using Autodesk.AutoCAD.DatabaseServices;
        //using Autodesk.AutoCAD.Geometry;
        //using Autodesk.AutoCAD.Runtime;
        //using System.Collections.Generic;

        //public class PolylineDirections
        //    {
        //        [CommandMethod("GetPolylinePointsByDirection")]
        //        public void GetPolylinePointsByDirection()
        //        {
        //            Document acDoc = Application.DocumentManager.MdiActiveDocument;
        //            Database acCurDb = acDoc.Database;

        //            // Start a transaction
        //            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        //            {
        //                // Create a PromptEntityOptions object to prompt the user to select a polyline
        //                PromptEntityOptions peo = new PromptEntityOptions("\nSelect a polyline: ");
        //                peo.SetRejectMessage("\nSelected entity is not a polyline.");
        //                peo.AddAllowedClass(typeof(Polyline), false);

        //                // Get the selected polyline
        //                PromptEntityResult per = acDoc.Editor.GetEntity(peo);
        //                if (per.Status != PromptStatus.OK)
        //                {
        //                    acTrans.Abort();
        //                    return;
        //                }

        //                // Open the selected polyline for read
        //                Polyline acPoly = acTrans.GetObject(per.ObjectId, OpenMode.ForRead) as Polyline;

        //                // Define lists to hold points in different directions
        //                List<Point2d> northPoints = new List<Point2d>();
        //                List<Point2d> southPoints = new List<Point2d>();
        //                List<Point2d> eastPoints = new List<Point2d>();
        //                List<Point2d> westPoints = new List<Point2d>();

        //                // Iterate through the polyline segments
        //                for (int i = 0; i < acPoly.NumberOfVertices; i++)
        //                {
        //                    if (acPoly.GetSegmentType(i) == SegmentType.Line)
        //                    {
        //                        LineSegment2d segment = acPoly.GetLineSegment2dAt(i);

        //                        // Determine the direction of the segment
        //                        if (segment.StartPoint.X == segment.EndPoint.X)
        //                        {
        //                            // Vertical line
        //                            if (segment.StartPoint.Y < segment.EndPoint.Y)
        //                            {
        //                                northPoints.Add(segment.StartPoint);
        //                                northPoints.Add(segment.EndPoint);
        //                            }
        //                            else
        //                            {
        //                                southPoints.Add(segment.StartPoint);
        //                                southPoints.Add(segment.EndPoint);
        //                            }
        //                        }
        //                        else if (segment.StartPoint.Y == segment.EndPoint.Y)
        //                        {
        //                            // Horizontal line
        //                            if (segment.StartPoint.X < segment.EndPoint.X)
        //                            {
        //                                eastPoints.Add(segment.StartPoint);
        //                                eastPoints.Add(segment.EndPoint);
        //                            }
        //                            else
        //                            {
        //                                westPoints.Add(segment.StartPoint);
        //                                westPoints.Add(segment.EndPoint);
        //                            }
        //                        }
        //                    }
        //                }

        //                // Output the results
        //                acDoc.Editor.WriteMessage($"\nNorth Points: {northPoints.Count}");
        //                acDoc.Editor.WriteMessage($"\nSouth Points: {southPoints.Count}");
        //                acDoc.Editor.WriteMessage($"\nEast Points: {eastPoints.Count}");
        //                acDoc.Editor.WriteMessage($"\nWest Points: {westPoints.Count}");

        //                // Commit the transaction
        //                acTrans.Commit();
        //            }
        //        }
        //    }

        #endregion

    }
}
