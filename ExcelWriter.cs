using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using Test.Properties;
using System.Threading.Tasks;

namespace Test
{
    public class ExcelWriter
    {
        public static async Task InitTable(List<Comment> comments)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo(@"/home/shastiva/Projects/Test/Test/Feedbacks.xlsx");

            await SaveExcelFile(comments, file);
        }

        private static async Task SaveExcelFile(List<Comment> persons, FileInfo file)
        {
            DeleteIfExists(file);

            using var package = new ExcelPackage(file);

            var ws = package.Workbook.Worksheets.Add("Information");

            var range = ws.Cells["A1"].LoadFromCollection(persons, true);
            range.AutoFitColumns();

            await package.SaveAsync();
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }
    }
}