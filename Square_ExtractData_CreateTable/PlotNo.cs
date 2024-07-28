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
    public class PlotNo
    {
        public string _PlotNo;
        public double _AreaInEast;
        public double _AreaInSouth;
        public double _AreaInWest;
        public double _AreaInNorth;
        public double _PlotArea;
        public double _MortgageArea;
        public double _UtilityArea;
        public double _AmenityArea;
        public string _EastInfo;
        public string _SouthInfo;
        public string _WestInfo;
        public string _NorthInfo;
        public Polyline _Polyline;
        public List<SurveyNo> _ParentSurveyNos = new List<SurveyNo>();
    }
}
