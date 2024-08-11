using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Square_ExtractData_CreateTable
{
    public static class Constants
    {
        public static string SurveyNoLayer = "_SurveyNo";
        public static string IndivPlotLayer = "_IndivSubPlot";
        //public static string IndivPlotDimLayer = "_IndivSubPlot_DIMENSION";
        public static string MortgageLayer = "_MortgageArea";
        public static string AmenityLayer = "_Amenity";
        //public static string AmenityDimLayer = "_Amenity_DIMENSION";
        public static string DocNoLayer = "_DocNo";
        public static string LandLordLayer = "_LandLord";
        public static string InternalRoadLayer = "_InternalRoad";

        public static string PlotLayer = "_Plot";
        public static string OpenSpaceLayer = "_OrganizedOpenSpace";
        public static string UtilityLayer = "_UtilityArea";
        public static string LeftOverOwnerLandLayer = "_LeftoverOwnersLand";
        public static string SideBoundaryLayer = "_SideBoundary";
        public static string MainRoadLayer = "_MainRoad";
        public static string SplayLayer = "_Splay";
        public static string RoadWideningLayer = "_RoadWidening";
        public static string GreenBufferZoneLayer = "_GreenBufferZone";

        public static string LWPOLYLINE = "LWPOLYLINE";
        public static string TEXT = "TEXT";
        public static string MTEXT = "MTEXT";

        public static int AreaDecimals = 2;
        public static int uniquePointsIdentifier = 1;
        public static double minArea = 1.0;

        public static string ExceltemplatePath = @"C:\Data\Square_Excel_Template.xlsx";
    }
}
