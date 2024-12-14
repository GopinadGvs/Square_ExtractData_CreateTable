using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Square_ExtractData_CreateTable
{
    public class AreaConstants
    {
        public double AmenityPercentage { get; set; }
        public double UtilityPercentage { get; set; }
        public double OpenSpacePercentage { get; set; }
        public double LayoutArea { get; set; }
        public double MortgagePercentage { get; set; }
        public double TotalPlotArea { get; set; }

        public AreaConstants(double layoutArea)
        {
            this.LayoutArea = layoutArea;

            if (layoutArea >= 50000)
            {
                AmenityPercentage = 3;
                UtilityPercentage = 1;                
            }
            else if (layoutArea < 50000)
            {
                AmenityPercentage = 2;
                UtilityPercentage = 0.5;
            }
            OpenSpacePercentage = 10;
            MortgagePercentage = 15;
        }

        public bool ValidateAmenity(double amenityArea)
        {
            double percentage = amenityArea / LayoutArea * 100;

            if (percentage >= AmenityPercentage)
                return true;

            return false;
        }

        public bool ValidateOpenSpace(double openSpaceArea)
        {
            double percentage = openSpaceArea / LayoutArea * 100;

            if (percentage >= OpenSpacePercentage)
                return true;

            return false;
        }

        public bool ValidateUtility(double utilityArea)
        {
            double percentage = utilityArea / LayoutArea * 100;

            if (percentage >= UtilityPercentage)
                return true;

            return false;
        }

        public bool ValidateMortgage(double mortgageArea)
        {
            double percentage = mortgageArea / TotalPlotArea * 100;

            if (percentage >= MortgagePercentage)
                return true;

            return false;
        }
    }
}
