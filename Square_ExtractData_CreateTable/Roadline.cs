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
    public class Roadline
    {
        //public List<string> _PlotNos = new List<string>();
        public string _RoadText = string.Empty;
        public Polyline _Polyline;
        public Point3dCollection _PolylinePoints = new Point3dCollection();
        //public List<SurveyNo> _ParentSurveyNos = new List<SurveyNo>();
        public Point3d Center;
        public ObjectId _RoadlineId;

    }
}
