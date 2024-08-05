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

        #region FillPointsByDirectionUsingCenterLogic

        //        using Autodesk.AutoCAD.ApplicationServices;
        //using Autodesk.AutoCAD.DatabaseServices;
        //using Autodesk.AutoCAD.EditorInput;
        //using Autodesk.AutoCAD.Geometry;
        //using Autodesk.AutoCAD.Runtime;
        //using System.Collections.Generic;
        //using System.Linq;

        //public class PolylineDirections
        //    {
        //        [CommandMethod("GetPolylinePointsByDirection")]
        //        public void GetPolylinePointsByDirection()
        //        {
        //            Document acDoc = Application.DocumentManager.MdiActiveDocument;
        //            Database acCurDb = acDoc.Database;

        //            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        //            {
        //                // Prompt the user to select a polyline
        //                PromptEntityOptions peo = new PromptEntityOptions("\nSelect a polyline: ");
        //                peo.SetRejectMessage("\nSelected entity is not a polyline.");
        //                peo.AddAllowedClass(typeof(Polyline), false);

        //                PromptEntityResult per = acDoc.Editor.GetEntity(peo);
        //                if (per.Status != PromptStatus.OK)
        //                {
        //                    return;
        //                }

        //                // Open the selected polyline for read
        //                Polyline acPoly = acTrans.GetObject(per.ObjectId, OpenMode.ForRead) as Polyline;

        //                // Get the center of the polyline
        //                Point3d polyCenter = GetPolylineCenter(acPoly);

        //                // Define lists to hold points in different directions
        //                List<Point3d> northPoints = new List<Point3d>();
        //                List<Point3d> southPoints = new List<Point3d>();
        //                List<Point3d> eastPoints = new List<Point3d>();
        //                List<Point3d> westPoints = new List<Point3d>();

        //                // Iterate through the polyline vertices
        //                for (int i = 0; i < acPoly.NumberOfVertices; i++)
        //                {
        //                    Point3d vertex = acPoly.GetPoint3dAt(i);

        //                    Vector3d direction = vertex - polyCenter;

        //                    if (direction.Y > 0)
        //                    {
        //                        northPoints.Add(vertex);
        //                    }
        //                    else if (direction.Y < 0)
        //                    {
        //                        southPoints.Add(vertex);
        //                    }

        //                    if (direction.X > 0)
        //                    {
        //                        eastPoints.Add(vertex);
        //                    }
        //                    else if (direction.X < 0)
        //                    {
        //                        westPoints.Add(vertex);
        //                    }
        //                }

        //                // Sort points by their distance to the center and take the two closest points
        //                northPoints = northPoints.OrderBy(p => p.DistanceTo(polyCenter)).Take(2).ToList();
        //                southPoints = southPoints.OrderBy(p => p.DistanceTo(polyCenter)).Take(2).ToList();
        //                eastPoints = eastPoints.OrderBy(p => p.DistanceTo(polyCenter)).Take(2).ToList();
        //                westPoints = westPoints.OrderBy(p => p.DistanceTo(polyCenter)).Take(2).ToList();

        //                // Output the results
        //                acDoc.Editor.WriteMessage("\nNorth Points: " + string.Join(", ", northPoints));
        //                acDoc.Editor.WriteMessage("\nSouth Points: " + string.Join(", ", southPoints));
        //                acDoc.Editor.WriteMessage("\nEast Points: " + string.Join(", ", eastPoints));
        //                acDoc.Editor.WriteMessage("\nWest Points: " + string.Join(", ", westPoints));

        //                // Commit the transaction
        //                acTrans.Commit();
        //            }
        //        }

        //        private Point3d GetPolylineCenter(Polyline poly)
        //        {
        //            Point3dCollection pts = new Point3dCollection();
        //            for (int i = 0; i < poly.NumberOfVertices; i++)
        //            {
        //                pts.Add(poly.GetPoint3dAt(i));
        //            }

        //            // Calculate the centroid
        //            double xSum = 0;
        //            double ySum = 0;
        //            double zSum = 0;

        //            foreach (Point3d pt in pts)
        //            {
        //                xSum += pt.X;
        //                ySum += pt.Y;
        //                zSum += pt.Z;
        //            }

        //            return new Point3d(xSum / pts.Count, ySum / pts.Count, zSum / pts.Count);
        //        }
        //    }

        #endregion


        #region Sum of area with points array as input

        //        using Autodesk.AutoCAD.ApplicationServices;
        //using Autodesk.AutoCAD.DatabaseServices;
        //using Autodesk.AutoCAD.Geometry;

        //public double CalculateSumOfAreas(Point3d[][] polygons)
        //    {
        //        double totalArea = 0.0;

        //        // Get the current document and database
        //        Document doc = Application.DocumentManager.MdiActiveDocument;
        //        Database db = doc.Database;

        //        // Start a transaction
        //        using (Transaction trans = db.TransactionManager.StartTransaction())
        //        {
        //            // Open the BlockTable for read
        //            BlockTable blockTable = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);

        //            foreach (Point3d[] points in polygons)
        //            {
        //                // Create a Polyline
        //                using (LWPOLYLINE polyline = new LWPOLYLINE())
        //                {
        //                    // Add points to the polyline
        //                    for (int i = 0; i < points.Length; i++)
        //                    {
        //                        polyline.AddVertexAt(i, new Point2d(points[i].X, points[i].Y), 0, 0, 0);
        //                    }
        //                    polyline.Closed = true; // Close the polyline to form a closed polygon

        //                    // Add the polyline to the current space
        //                    BlockTableRecord btr = (BlockTableRecord)trans.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
        //                    btr.AppendEntity(polyline);
        //                    trans.AddNewlyCreatedDBObject(polyline, true);

        //                    // Calculate the area of the polyline
        //                    double area = polyline.Area;
        //                    totalArea += area;
        //                }
        //            }

        //            // Commit the transaction
        //            trans.Commit();
        //        }

        //        return totalArea;
        //    }

        #endregion

        #region Check whether point is inside polyline or not

        //        using Autodesk.AutoCAD.ApplicationServices;
        //using Autodesk.AutoCAD.DatabaseServices;
        //using Autodesk.AutoCAD.Geometry;
        //using Autodesk.AutoCAD.Runtime;

        //public bool IsPointInsidePolyline(Point3d point, Polyline polyline)
        //    {
        //        // Convert Point3d to Point2d
        //        Point2d pt2d = new Point2d(point.X, point.Y);

        //        // Check if the polyline is closed and if the point is inside it
        //        if (polyline.Closed)
        //        {
        //            return polyline.Contains(pt2d);
        //        }
        //        else
        //        {
        //            return false; // Open polylines don't define closed areas
        //        }
        //    }

        //    [CommandMethod("TestPointInPolyline")]
        //    public void TestPointInPolyline()
        //    {
        //        // Get the current document and database
        //        Document doc = Application.DocumentManager.MdiActiveDocument;
        //        Database db = doc.Database;

        //        // Start a transaction
        //        using (Transaction trans = db.TransactionManager.StartTransaction())
        //        {
        //            // Open the BlockTable for read
        //            BlockTable blockTable = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);

        //            // Open the ModelSpace for read
        //            BlockTableRecord modelSpace = (BlockTableRecord)trans.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

        //            // Example point to test
        //            Point3d testPoint = new Point3d(10, 10, 0);

        //            // Iterate through entities in ModelSpace
        //            foreach (ObjectId objId in modelSpace)
        //            {
        //                DBObject dbObj = trans.GetObject(objId, OpenMode.ForRead);
        //                if (dbObj is Polyline polyline)
        //                {
        //                    // Check if the point is inside the polyline
        //                    bool isInside = IsPointInsidePolyline(testPoint, polyline);
        //                    doc.Editor.WriteMessage($"\nPoint {testPoint} is inside polyline: {isInside}");
        //                }
        //            }

        //            // Commit the transaction
        //            trans.Commit();
        //        }
        //    }


        #endregion

        #region sum of area with point array as input, sort them to form a proper closed polygon

        //        using Autodesk.AutoCAD.ApplicationServices;
        //using Autodesk.AutoCAD.DatabaseServices;
        //using Autodesk.AutoCAD.Geometry;
        //using Autodesk.AutoCAD.Runtime;
        //using System;
        //using System.Collections.Generic;
        //using System.Linq;

        //public class PolylineAreaCalculator
        //    {
        //        public double CalculateAreaFromPoints(Point3d[] points)
        //        {
        //            // Sort the points to form a proper closed polyline
        //            List<Point3d> sortedPoints = SortPoints(points);

        //            // Create a closed polyline from the sorted points
        //            Polyline polyline = CreateClosedPolyline(sortedPoints);

        //            // Calculate the area of the polyline
        //            double area = polyline.Area;

        //            return area;
        //        }

        //        private List<Point3d> SortPoints(Point3d[] points)
        //        {
        //            // This is a simple example and may not handle all cases
        //            // For complex cases, consider using computational geometry libraries
        //            // For now, we'll assume points are approximately ordered correctly

        //            // Convert to Point2d for sorting
        //            List<Point2d> point2dList = points.Select(p => new Point2d(p.X, p.Y)).ToList();

        //            // Perform a sort based on polar angle from centroid
        //            Point2d centroid = new Point2d(point2dList.Average(p => p.X), point2dList.Average(p => p.Y));
        //            point2dList.Sort((p1, p2) => ComparePolarAngles(p1, p2, centroid));

        //            // Convert back to Point3d
        //            List<Point3d> sortedPoints = point2dList.Select(p => new Point3d(p.X, p.Y, 0)).ToList();

        //            return sortedPoints;
        //        }

        //        private int ComparePolarAngles(Point2d p1, Point2d p2, Point2d centroid)
        //        {
        //            double angle1 = Math.Atan2(p1.Y - centroid.Y, p1.X - centroid.X);
        //            double angle2 = Math.Atan2(p2.Y - centroid.Y, p2.X - centroid.X);
        //            return angle1.CompareTo(angle2);
        //        }

        //        private Polyline CreateClosedPolyline(List<Point3d> points)
        //        {
        //            Polyline polyline = new Polyline();

        //            // Add points to polyline
        //            for (int i = 0; i < points.Count; i++)
        //            {
        //                Point2d pt2d = new Point2d(points[i].X, points[i].Y);
        //                polyline.AddVertexAt(i, pt2d, 0, 0, 0);
        //            }

        //            polyline.Closed = true; // Close the polyline

        //            return polyline;
        //        }

        //        [CommandMethod("CalculateAreaFromPoints")]
        //        public void CalculateAreaFromPointsCommand()
        //        {
        //            // Example points
        //            Point3d[] points = new Point3d[]
        //            {
        //            new Point3d(0, 0, 0),
        //            new Point3d(10, 0, 0),
        //            new Point3d(10, 10, 0),
        //            new Point3d(0, 10, 0)
        //            };

        //            double area = CalculateAreaFromPoints(points);
        //            Application.ShowAlertDialog($"Calculated Area: {area}");
        //        }
        //    }


        #endregion

        #region Angle of a line

        //        using Autodesk.AutoCAD.ApplicationServices;
        //using Autodesk.AutoCAD.DatabaseServices;
        //using Autodesk.AutoCAD.Geometry;
        //using Autodesk.AutoCAD.Runtime;
        //using System;

        //public class LineAngleCalculator
        //    {
        //        public double CalculateLineAngle(Line line)
        //        {
        //            // Get the direction vector of the line
        //            Vector3d direction = line.Delta;

        //            // Convert the direction vector to 2D by ignoring the Z component
        //            double x = direction.X;
        //            double y = direction.Y;

        //            // Calculate the angle between the direction vector and the X-axis
        //            double angleRadians = Math.Atan2(y, x);

        //            // Convert the angle from radians to degrees
        //            double angleDegrees = angleRadians * (180.0 / Math.PI);

        //            // Normalize the angle to be in the range [0, 360) degrees
        //            if (angleDegrees < 0)
        //            {
        //                angleDegrees += 360;
        //            }

        //            return angleDegrees;
        //        }

        //        [CommandMethod("GetLineAngle")]
        //        public void GetLineAngleCommand()
        //        {
        //            Document doc = Application.DocumentManager.MdiActiveDocument;
        //            Database db = doc.Database;
        //            Editor ed = doc.Editor;

        //            // Prompt user to select a line
        //            PromptEntityOptions opt = new PromptEntityOptions("\nSelect a line: ");
        //            opt.SetRejectMessage("\nEntity must be a line.");
        //            opt.AddAllowedClass(typeof(Line), false);

        //            PromptEntityResult res = ed.GetEntity(opt);
        //            if (res.Status != PromptStatus.OK)
        //                return;

        //            using (Transaction trans = db.TransactionManager.StartTransaction())
        //            {
        //                Line line = trans.GetObject(res.ObjectId, OpenMode.ForRead) as Line;

        //                if (line != null)
        //                {
        //                    double angle = CalculateLineAngle(line);
        //                    ed.WriteMessage($"\nThe angle of the line is {angle:F2} degrees.");
        //                }

        //                trans.Commit();
        //            }
        //        }
        //    }




        #endregion

    }
}
