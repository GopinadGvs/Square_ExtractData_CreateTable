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
    public class SurveyNo
    {
        public List<Plot> _PlotNos = new List<Plot>();
        public List<Amenity> _AmenityPlots = new List<Amenity>();

        public List<Mortgage> _MortgagePlotNos = new List<Mortgage>();

        public string _SurveyNo = string.Empty;
        public Polyline _Polyline;
        public Point3dCollection _PolylinePoints = new Point3dCollection();
        public Point3d Center;

        //public List<Point2d> northPoints = new List<Point2d>();
        //public List<Point2d> southPoints = new List<Point2d>();
        //public List<Point2d> eastPoints = new List<Point2d>();
        //public List<Point2d> westPoints = new List<Point2d>();

        public List<Point3d> eastPoints = new List<Point3d>();
        public List<Point3d> southPoints = new List<Point3d>();
        public List<Point3d> westPoints = new List<Point3d>();
        public List<Point3d> northPoints = new List<Point3d>();
    }
}
