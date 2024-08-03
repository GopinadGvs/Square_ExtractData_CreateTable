using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace Square_ExtractData_CreateTable
{
    public class Plot
    {
        public Plot()
        {
            IsAmenity = false;
        }
        public string _PlotNo;
        public List<SDimension> _SizesInEast = new List<SDimension>();
        public List<SDimension> _SizesInSouth = new List<SDimension>();
        public List<SDimension> _SizesInWest = new List<SDimension>();
        public List<SDimension> _SizesInNorth = new List<SDimension>();
        public double _Area;
        public double _PlotArea;
        public double _MortgageArea;
        public double _AmenityArea;
        public string _EastInfo;
        public string _SouthInfo;
        public string _WestInfo;
        public string _NorthInfo;
        public Polyline _Polyline;
        public List<SurveyNo> _ParentSurveyNos = new List<SurveyNo>();
        public bool IsPlotArea;
        public bool IsMortgageArea;
        public bool IsAmenity;
        public Point3dCollection _PolylinePoints = new Point3dCollection();
        public Point3d Center;
        public List<Point3d> eastPoints = new List<Point3d>();
        public List<Point3d> southPoints = new List<Point3d>();
        public List<Point3d> westPoints = new List<Point3d>();
        public List<Point3d> northPoints = new List<Point3d>();

        public List<LineSegment3d> eastLineSegment = new List<LineSegment3d>();
        public List<LineSegment3d> southLineSegment = new List<LineSegment3d>();
        public List<LineSegment3d> westLineSegment = new List<LineSegment3d>();
        public List<LineSegment3d> northLineSegment = new List<LineSegment3d>();


        public List<Point3d> MortgagePoints = new List<Point3d>();

        public List<SDimension> _AllDims = new List<SDimension>();
    }
}
