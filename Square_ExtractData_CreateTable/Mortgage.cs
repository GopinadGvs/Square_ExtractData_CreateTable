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
    public class Mortgage
    {
        public List<Plot> _PlotNos = new List<Plot>();
        public string _SurveyNo = string.Empty;
        public Polyline _Polyline;
        public Point3dCollection _PolylinePoints = new Point3dCollection();
    }
}
