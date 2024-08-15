using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Square_ExtractData_CreateTable
{
    public static class FindMissingNumber
    {
        #region AutoLisp working Logic

        //public class CustomStringComparer : IComparer<string>
        //{
        //    public int Compare(string x, string y)
        //    {
        //        if (AnyParamIsNull(x, y, out var result))
        //        {
        //            return result.Value;
        //        }
        //        SplitInput(x, out var alphabets, out var number);
        //        SplitInput(y, out var alphabets2, out var number2);
        //        if (alphabets.Length == alphabets2.Length)
        //        {
        //            result = string.Compare(alphabets, alphabets2);
        //            if (result.Value == 0)
        //            {
        //                if (number < number2)
        //                {
        //                    return -1;
        //                }
        //                if (number > number2)
        //                {
        //                    return 1;
        //                }
        //                return 0;
        //            }
        //            return result.Value;
        //        }
        //        return alphabets.Length - alphabets2.Length;
        //    }

        //    private bool SplitInput(string input, out string alphabets, out double number)
        //    {
        //        Regex regex = new Regex("\\d*[.]?\\d+");
        //        Match match = regex.Match(input);
        //        if (match.Success)
        //        {
        //            number = double.Parse(match.Value, CultureInfo.InvariantCulture);
        //            alphabets = input.Replace(match.Value, "");
        //            return true;
        //        }
        //        throw new ArgumentException("Input string is not of correct format", input);
        //    }

        //    private bool AnyParamIsNull(string x, string y, out int? result)
        //    {
        //        result = null;
        //        if (x == null)
        //        {
        //            if (y == null)
        //            {
        //                result = 0;
        //                return true;
        //            }
        //            result = -1;
        //            return true;
        //        }
        //        if (y == null)
        //        {
        //            result = 1;
        //            return true;
        //        }
        //        return false;
        //    }
        //}

        //private static void Main(string[] args)
        //{
        //    int num = 0;
        //    int num2 = 0;
        //    string[] array = File.ReadAllLines("C:\\Data\\fmi.csv");
        //    //reverse the array items
        //    for (num = 0; num < array.Length; num++)
        //    {
        //        string[] array2 = array[num].Split(',');
        //        for (num2 = 0; num2 < array2.Length - 1; num2++)
        //        {
        //            string text = array2[num2].ToString();
        //            string text2 = array2[num2 + 1].ToString();
        //            array[num] = text2 + "," + text;
        //        }
        //    }
        //    for (int i = 0; i < array.Length; i++)
        //    {
        //    }
        //    List<string> list = new List<string>(array);
        //    //sort the items
        //    list.Sort(new CustomStringComparer());
        //    //reverse back to original state
        //    for (int j = 0; j < list.Count; j++)
        //    {
        //        string[] array3 = list[j].Split(',');
        //        for (num2 = 0; num2 < array3.Length - 1; num2++)
        //        {
        //            string text = array3[num2].ToString();
        //            string text2 = array3[num2 + 1].ToString();
        //            list[j] = text2 + "," + text;
        //        }
        //    }
        //    for (int k = 0; k < list.Count; k++)
        //    {
        //    }
        //    //write to csv again
        //    File.WriteAllLines("C:\\Data\\fmi.csv", list);
        //    string[] array4 = File.ReadAllLines("C:\\Data\\fmi.csv");
        //    StreamWriter streamWriter = new StreamWriter("C:\\Data\\fmi.csv");
        //    StreamWriter streamWriter2 = new StreamWriter("C:\\Data\\fms.csv");
        //    string[] array5 = array4;
        //    //write numbers to fmi csv and string type to fms csv file
        //    foreach (string text3 in array5)
        //    {
        //        string[] array6 = text3.Split(',');
        //        if (int.TryParse(array6[0], out var _))
        //        {
        //            streamWriter.WriteLine(array6[0]);
        //        }
        //        else
        //        {
        //            streamWriter2.WriteLine(array6[0]);
        //        }
        //    }
        //    streamWriter.Close();
        //    streamWriter2.Close();
        //    string[] source = File.ReadAllLines("C:\\Data\\fmi.csv");
        //    List<int> list2 = source.Select((string s) => Convert.ToInt32(s)).ToList();
        //    int num3 = list2.Min();
        //    int num4 = list2.Max();
        //    //logic to get missing numbers from range
        //    IEnumerable<int> enumerable = Enumerable.Range(num3, num4 - num3).Except(list2);
        //    StreamWriter streamWriter3 = new StreamWriter("C:\\Data\\fmm.csv");
        //    //write missing numbers to fmm csv
        //    foreach (int item in enumerable)
        //    {
        //        streamWriter3.WriteLine(item);
        //    }
        //    streamWriter3.Close();
        //    //create a dictionary with item number, it's repetative count
        //    Dictionary<int, int> dictionary = new Dictionary<int, int>();
        //    foreach (int item2 in list2)
        //    {
        //        if (dictionary.ContainsKey(item2))
        //        {
        //            dictionary[item2]++;
        //        }
        //        else
        //        {
        //            dictionary[item2] = 1;
        //        }
        //    }
        //    //write duplicates to fdup file
        //    StreamWriter streamWriter4 = new StreamWriter("C:\\Data\\fdup.csv");
        //    foreach (KeyValuePair<int, int> item3 in dictionary)
        //    {
        //        if (item3.Value > 1)
        //        {
        //            streamWriter4.WriteLine("Value {0} occurred {1} times,", item3.Key, item3.Value);
        //        }
        //    }
        //    streamWriter4.Close();
        //    StreamWriter streamWriter5 = new StreamWriter("C:\\Data\\datafmi.txt");
        //    streamWriter5.WriteLine("yes");
        //    streamWriter5.WriteLine(num3);
        //    streamWriter5.WriteLine(num4);
        //    streamWriter5.Close();
        //}

        #endregion


        public static string missingNumbersString;
        public static List<string> duplicateNumbersString;
        public static string otherNumbersString;
    }
}
