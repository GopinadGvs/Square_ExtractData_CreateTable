using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using File = System.IO.File;
using Path = System.IO.Path;

namespace Square_ExtractData_CreateTable
{
    public static class LogWriter
    {
        private static string _mFilepath = @"C:\users\Public\Documents\SquareLog.txt";
        public static void LogWrite(string logMessage)
        {
            if (!File.Exists(_mFilepath))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_mFilepath));
                File.Create(_mFilepath).Dispose();
            }
            using (StreamWriter w = File.AppendText(_mFilepath))
            {
                Log(logMessage, w);
                
            }
        }

        static void Log(string logMessage, TextWriter txtWriter)
        {
            txtWriter.Write("\r\nLog Entry : ");
            txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
                DateTime.Now.ToLongDateString());
            txtWriter.WriteLine("  :");
            txtWriter.WriteLine("  :{0}", logMessage);
            txtWriter.WriteLine("-------------------------------");
        }

      
    }
}
