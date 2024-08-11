using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;


namespace Square_ExtractData_CreateTable
{
    #region string comparer reference

    //class Program
    //{
    //    static void Main()
    //    {
    //        List<string> input = new List<string> { "1", "10", "11", "2", "20", "20A", "20B", "3", "30", "31C", "4" };

    //        input.Sort(new AlphanumericComparer());

    //        foreach (string s in input)
    //        {
    //            Console.WriteLine(s);
    //        }
    //    }
    //}

    //public class AlphanumericComparer : IComparer<string>
    //{
    //    public int Compare(string x, string y)
    //    {
    //        if (x == y) return 0;

    //        // Split both strings into numeric and non-numeric parts
    //        var xParts = SplitNumericAndNonNumericParts(x);
    //        var yParts = SplitNumericAndNonNumericParts(y);

    //        int minLength = Math.Min(xParts.Count, yParts.Count);
    //        for (int i = 0; i < minLength; i++)
    //        {
    //            int comparisonResult = ComparePart(xParts[i], yParts[i]);
    //            if (comparisonResult != 0)
    //            {
    //                return comparisonResult;
    //            }
    //        }

    //        // If one string is a prefix of the other, the shorter string should come first
    //        return xParts.Count.CompareTo(yParts.Count);
    //    }

    //    private List<string> SplitNumericAndNonNumericParts(string input)
    //    {
    //        var parts = new List<string>();
    //        var matches = Regex.Matches(input, @"(\d+|\D+)");
    //        foreach (Match match in matches)
    //        {
    //            parts.Add(match.Value);
    //        }
    //        return parts;
    //    }

    //    private int ComparePart(string xPart, string yPart)
    //    {
    //        // If both parts are numeric, compare them as integers
    //        if (int.TryParse(xPart, out int xNumber) && int.TryParse(yPart, out int yNumber))
    //        {
    //            return xNumber.CompareTo(yNumber);
    //        }

    //        // Otherwise, compare them as strings
    //        return string.Compare(xPart, yPart, StringComparison.Ordinal);
    //    }
    //}

    #endregion

    public class AlphanumericPlotComparer : IComparer<Plot>
    {
        public int Compare(Plot x, Plot y)
        {
            if (x == null || y == null)
                return 0;

            return AlphanumericCompare(x._PlotNo, y._PlotNo);
        }

        private int AlphanumericCompare(string x, string y)
        {
            if (x == y) return 0;

            // Split both strings into numeric and non-numeric parts
            var xParts = SplitNumericAndNonNumericParts(x);
            var yParts = SplitNumericAndNonNumericParts(y);

            int minLength = Math.Min(xParts.Count, yParts.Count);
            for (int i = 0; i < minLength; i++)
            {
                int comparisonResult = ComparePart(xParts[i], yParts[i]);
                if (comparisonResult != 0)
                {
                    return comparisonResult;
                }
            }

            // If one string is a prefix of the other, the shorter string should come first
            return xParts.Count.CompareTo(yParts.Count);
        }

        private List<string> SplitNumericAndNonNumericParts(string input)
        {
            var parts = new List<string>();
            var matches = Regex.Matches(input, @"(\d+|\D+)");
            foreach (Match match in matches)
            {
                parts.Add(match.Value);
            }
            return parts;
        }

        private int ComparePart(string xPart, string yPart)
        {
            // If both parts are numeric, compare them as integers
            if (int.TryParse(xPart, out int xNumber) && int.TryParse(yPart, out int yNumber))
            {
                return xNumber.CompareTo(yNumber);
            }

            // Otherwise, compare them as strings
            return string.Compare(xPart, yPart, StringComparison.Ordinal);
        }

    }
}
