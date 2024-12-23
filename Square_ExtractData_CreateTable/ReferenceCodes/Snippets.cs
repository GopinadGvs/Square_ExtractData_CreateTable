﻿using System;
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

        #region Extract Number from string

        //        using System;
        //using System.Collections.Generic;
        //using System.Text.RegularExpressions;

        //namespace QuickConsole
        //    {
        //        class Program
        //        {
        //            static void Main()
        //            {
        //                // Sample strings
        //                string[] inputs = { "12Meters road", "12 Mts. Road", "6 feet road" };

        //                foreach (var input in inputs)
        //                {
        //                    // Extract numbers from the string
        //                    List<string> numbers = ExtractNumbers(input);
        //                    Console.WriteLine($"Input: {input}");
        //                    Console.WriteLine($"Extracted Numbers: {string.Join(", ", numbers)}");
        //                    Console.WriteLine();
        //                }
        //            }

        //            static List<string> ExtractNumbers(string input)
        //            {
        //                // Regular expression to match numbers (including decimals)
        //                Regex regex = new Regex(@"\d+(\.\d+)?");

        //                // Find all matches
        //                MatchCollection matches = regex.Matches(input);

        //                // Collect all matches into a list of strings
        //                List<string> numbers = new List<string>();
        //                foreach (Match match in matches)
        //                {
        //                    numbers.Add(match.Value);
        //                }

        //                return numbers;
        //            }
        //        }
        //    }


        #endregion

        #region Check and create a Layer with color
        //        using Autodesk.AutoCAD.ApplicationServices;
        //using Autodesk.AutoCAD.DatabaseServices;
        //using Autodesk.AutoCAD.Runtime;
        //using Autodesk.AutoCAD.Colors;

        //public class LayerManager
        //    {
        //        [CommandMethod("CheckOrCreateLayer")]
        //        public void CheckOrCreateLayer()
        //        {
        //            // Get the current database and transaction manager
        //            Database db = HostApplicationServices.WorkingDatabase;
        //            TransactionManager tm = db.TransactionManager;

        //            using (Transaction tr = tm.StartTransaction())
        //            {
        //                // Open the LayerTable for read
        //                LayerTable layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

        //                // Check if the "_DocNo" layer exists
        //                bool layerExists = layerTable.Has("_DocNo");

        //                if (!layerExists)
        //                {
        //                    // Create a new layer table record for "_DocNo"
        //                    LayerTableRecord layerTableRecord = new LayerTableRecord
        //                    {
        //                        Name = "_DocNo",
        //                        Color = Color.FromRgb(255, 0, 0) // Red color
        //                    };

        //                    // Upgrade the LayerTable to write
        //                    layerTable.UpgradeOpen();

        //                    // Add the new layer to the LayerTable
        //                    layerTable.Add(layerTableRecord);

        //                    // Add the new LayerTableRecord to the transaction
        //                    tr.AddNewlyCreatedDBObject(layerTableRecord, true);

        //                    // Commit the transaction
        //                    tr.Commit();

        //                    // Inform the user
        //                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nLayer '_DocNo' has been created with red color.");
        //                }
        //                else
        //                {
        //                    // Inform the user
        //                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nLayer '_DocNo' already exists.");
        //                }
        //            }
        //        }
        //    }

        #endregion


        #region User Prompts

        //        using Autodesk.AutoCAD.ApplicationServices;
        //using Autodesk.AutoCAD.DatabaseServices;
        //using Autodesk.AutoCAD.Runtime;
        //using Autodesk.AutoCAD.EditorInput;

        //public class CopyPolylineToLayer
        //    {
        //        [CommandMethod("CopyPolylineToLayer")]
        //        public void CopyPolylineToAnotherLayer()
        //        {
        //            Document doc = Application.DocumentManager.MdiActiveDocument;
        //            Database db = doc.Database;
        //            Editor ed = doc.Editor;

        //            // Prompt the user to select a polyline
        //            PromptEntityOptions peo = new PromptEntityOptions("\nSelect a polyline to copy: ");
        //            peo.SetRejectMessage("\nSelected entity is not a polyline.");
        //            peo.AddAllowedClass(typeof(Polyline), true);

        //            PromptEntityResult per = ed.GetEntity(peo);

        //            if (per.Status != PromptStatus.OK)
        //            {
        //                ed.WriteMessage("\nCommand canceled.");
        //                return;
        //            }

        //            using (Transaction tr = db.TransactionManager.StartTransaction())
        //            {
        //                // Open the selected polyline for read
        //                Polyline polyline = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Polyline;

        //                if (polyline != null)
        //                {
        //                    // Clone the polyline
        //                    Polyline clonedPolyline = polyline.Clone() as Polyline;

        //                    // Prompt the user for the target layer name
        //                    PromptStringOptions pso = new PromptStringOptions("\nEnter target layer name: ");
        //                    pso.AllowSpaces = true;
        //                    PromptResult pr = ed.GetString(pso);

        //                    if (pr.Status == PromptStatus.OK)
        //                    {
        //                        string targetLayerName = pr.StringResult;

        //                        // Check if the layer exists, if not, create it
        //                        LayerTable layerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

        //                        if (!layerTable.Has(targetLayerName))
        //                        {
        //                            using (LayerTableRecord newLayer = new LayerTableRecord())
        //                            {
        //                                newLayer.Name = targetLayerName;
        //                                layerTable.UpgradeOpen();
        //                                layerTable.Add(newLayer);
        //                                tr.AddNewlyCreatedDBObject(newLayer, true);
        //                            }
        //                        }

        //                        // Set the cloned polyline to the target layer
        //                        clonedPolyline.Layer = targetLayerName;

        //                        // Add the cloned polyline to the model space
        //                        BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        //                        BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

        //                        modelSpace.AppendEntity(clonedPolyline);
        //                        tr.AddNewlyCreatedDBObject(clonedPolyline, true);

        //                        ed.WriteMessage($"\nPolyline copied to layer '{targetLayerName}'.");
        //                    }
        //                    else
        //                    {
        //                        ed.WriteMessage("\nNo layer name provided. Command canceled.");
        //                    }
        //                }
        //                else
        //                {
        //                    ed.WriteMessage("\nSelected entity is not a polyline.");
        //                }

        //                // Commit the transaction
        //                tr.Commit();
        //            }
        //        }
        //    }

        #endregion

        #region Explode polyline and create region

        //        using Autodesk.AutoCAD.ApplicationServices;
        //using Autodesk.AutoCAD.DatabaseServices;
        //using Autodesk.AutoCAD.Geometry;
        //using Autodesk.AutoCAD.Runtime;

        //public class PolylineContainment
        //    {
        //        [CommandMethod("CheckPolylineContainment")]
        //        public void CheckPolylineContainment()
        //        {
        //            Document doc = Application.DocumentManager.MdiActiveDocument;
        //            Database db = doc.Database;

        //            using (Transaction tr = db.TransactionManager.StartTransaction())
        //            {
        //                // Assuming we have the ObjectIds of the two polylines
        //                ObjectId outerPolylineId = new ObjectId(); // Set this to the ObjectId of the outer polyline
        //                ObjectId innerPolylineId = new ObjectId(); // Set this to the ObjectId of the inner polyline

        //                Polyline outerPolyline = tr.GetObject(outerPolylineId, OpenMode.ForRead) as Polyline;
        //                Polyline innerPolyline = tr.GetObject(innerPolylineId, OpenMode.ForRead) as Polyline;

        //                if (outerPolyline != null && innerPolyline != null)
        //                {
        //                    bool isInside = IsPolylineInside(outerPolyline, innerPolyline);

        //                    if (isInside)
        //                    {
        //                        doc.Editor.WriteMessage("\nThe inner polyline is inside the outer polyline.");
        //                    }
        //                    else
        //                    {
        //                        doc.Editor.WriteMessage("\nThe inner polyline is not inside the outer polyline.");
        //                    }
        //                }

        //                tr.Commit();
        //            }
        //        }

        //        private bool IsPolylineInside(Polyline outerPolyline, Polyline innerPolyline)
        //        {
        //            // Loop through each vertex of the inner polyline
        //            for (int i = 0; i < innerPolyline.NumberOfVertices; i++)
        //            {
        //                Point3d vertex = innerPolyline.GetPoint3dAt(i);

        //                // Check if the point is inside the outer polyline
        //                if (!IsPointInsidePolyline(outerPolyline, vertex))
        //                {
        //                    return false; // If any point is outside, the polyline is not inside
        //                }
        //            }
        //            return true;
        //        }

        //        private bool IsPointInsidePolyline(Polyline polyline, Point3d point)
        //        {
        //            // Convert polyline to Region
        //            using (var db = polyline.Database)
        //            {
        //                using (var region = new Region())
        //                {
        //                    DBObjectCollection dbObjCollection = new DBObjectCollection();
        //                    polyline.Explode(dbObjCollection);

        //                    DBObjectCollection regionColl = new DBObjectCollection();
        //                    Region.CreateFromCurves(dbObjCollection, regionColl);

        //                    Region regionEntity = regionColl[0] as Region;

        //                    if (regionEntity != null)
        //                    {
        //                        return regionEntity.PointContainment(new Point2d(point.X, point.Y)) == PointContainment.Inside;
        //                    }
        //                }
        //            }
        //            return false;
        //        }
        //    }


        #endregion

    }
}
