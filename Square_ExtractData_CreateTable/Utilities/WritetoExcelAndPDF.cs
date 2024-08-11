using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Library_NA_NA_NA_GenerateExcelandPDFReport;

namespace Square_ExtractData_CreateTable
{
    public static class WritetoExcelAndPDF
    {
        public static void WritetoExcel2(string prefix, string path, DataTable dt1)
        {
            VctDataTableRepository repo = new VctDataTableRepository();
            repo.TemplatePath = @"C:\Data\Square_Excel_Template.xlsx";
            repo.Prefix = prefix;
            repo.SavePath = path;
            if (!Directory.Exists(repo.SavePath))
                Directory.CreateDirectory(repo.SavePath);
            repo.MergePdf = true;
            repo.OpenExcelAfterGenerate = true;
            repo.OpenPdfAfterGenerate = true;

            //dt1.Rows.Add(new object[] { "James Bond, LLC", 120, "Garrison" });
            //dt1.Rows.Add(new object[] { "LLC", 10, "Gar" });
            //dt1.Rows.Add(new object[] { "Bond, LLC", 10, "Gar" });

            var highlighter = new StyleSettings()
            {
                BackColor = new HighlightColor() { B = 144, G = 238, R = 144 },
                WordWrap = true
            };

            VctDataTable dataTable1 = new VctDataTable(dt1)
            {
                SheetName = "Meters",
                PrintHeader = false,
                StartRow = 6
            };

            //dataTable1.Rows[0].Cells[0].UpdateSettings = true;
            //dataTable1.Rows[0].Cells[0].Settings = highlighter;
            //dataTable1.Rows[2].Cells[2].UpdateSettings = true;
            //dataTable1.Rows[2].Cells[2].Settings = highlighter;

            repo.VctDataTables.Add(dataTable1);
            repo.GenerateReport();
        }
    }
}
