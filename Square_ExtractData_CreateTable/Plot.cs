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
        public string _PlotNo;
        public double _SizeInEast;
        public double _SizeInSouth;
        public double _SizeInWest;
        public double _SizeInNorth;
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
        public bool IsUtilityArea;
        public Point3dCollection _PolylinePoints = new Point3dCollection();
        public Point3d Center;
        public List<Point2d> northPoints = new List<Point2d>();
        public List<Point2d> southPoints = new List<Point2d>();
        public List<Point2d> eastPoints = new List<Point2d>();
        public List<Point2d> westPoints = new List<Point2d>();

    }
}
