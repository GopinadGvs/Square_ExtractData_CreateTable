using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Square_ExtractData_CreateTable
{
    public static class Constants
    {
        //public static string SurveyNoLayer = "_SurveyNo";

        //survey no layer name updated to handle landlord sub 
        public static string SurveyNoLayer = "_LandLord_Sub";

        public static LayersLayer Plot;
        public static LayersLayer SurveyNo;
        public static LayersLayer IndivPlot;
        public static LayersLayer Mortgage;
        public static LayersLayer Amenity;
        public static LayersLayer Utility;
        public static LayersLayer InternalRoad;
        public static LayersLayer MainRoad;
        public static LayersLayer OrganizedOpenSpace;
        public static LayersLayer LeftoverOwnersLand;
        public static LayersLayer Splay;
        public static LayersLayer RoadWidening;
        public static LayersLayer GreenBufferZone;
        public static LayersLayer DocNo;
        public static LayersLayer LandLord;
        public static LayersLayer LandLord_Sub;
        public static LayersLayer Extent;
        public static LayersLayer SideBoundary;
        public static LayersLayer FreeSpace;
        public static LayersLayer FreeSpace_NameCheck;
        public static LayersLayer FreeSpace_DocCheck;
        public static LayersLayer FreeSpace_ExtentCheck;

        public static List<LayersLayer> LayersList = new List<LayersLayer>();

        public static string SurveyNoMainLayer = "_SurveyNo";

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

        public static string LandLordSubLayer = "_LandLord_Sub";

        public static string FreeSpaceLayer = "_FreeSpace";


        public static string LWPOLYLINE = "LWPOLYLINE";
        public static string TEXT = "TEXT";
        public static string MTEXT = "MTEXT";

        public static int AreaDecimals = 2;
        public static int uniquePointsIdentifier = 1;
        public static int uniquePointsIdentifierNew = 2;
        public static double minArea = 0.1;
        public static double areaTolerance = 0.01;


        public static string ExceltemplatePath = @"C:\Data\Square_Excel_Template.xlsx";

        public static string MExportHelpPDF = @"C:\Data\LayerSettings_V1.2.0.0.pdf";

        public static string ConfigPath = @"C:\Data\Config.xml";
    }
}
