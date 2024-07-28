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
        public List<PlotNo> _PlotNos = new List<PlotNo>();
        public string _SurveyNo = string.Empty;
        //public Dictionary<string, Polyline> _SurveyNoVsPolylines;
        public Polyline _Polyline;
    }
}
