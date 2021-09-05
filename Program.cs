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
    class Program
    {
        static async Task Main(string[] args)
        {
            var page = 1;
            var check = true;
            
            var url =
                "https://by-napi.wildberries.ru" +
                "/api/product/feedback" +
                "/5736970?brandId=4246" +
                $"&page={page}&order=Desc&field=Date" +
                "&withPhoto=False&_app-type=sitemobile";
            List<Person> persons = new List<Person>();

            while (check)
            {
                try
                {
                    var response = JObject.Parse(GetResponse(url));
                    var comments = response["data"]?["feedback"];

                    if ((int) comments["state"] == -1)
                    {
                        check = false;
                        break;
                    }

                    for (int i = 0; i < comments.Count(); i++)
                    {
                        Person pers = new Person();

                        pers.Name = (string) comments[i]["userName"];
                        pers.Date = (string) comments[i]["date"];
                        pers.Text = (string) comments[i]["text"];
                        pers.Raiting = (int) comments[i]["mark"];
                        pers.Like = (int) comments[i]["votesUp"];
                        pers.Dislike = (int) comments[i]["votesDown"];

                        persons.Add(pers);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                page++;
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo(@"/home/shastiva/Projects/Test/Test/Feedbacks.xlsx");

            await SaveExcelFile(persons, file);
        }

        private static async Task SaveExcelFile(List<Person> persons, FileInfo file)
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

        public static String GetResponse(string url)
        {
            string htmlCode;
            using (WebClient client = new WebClient())
            {
                client.Encoding = Encoding.UTF8;
                htmlCode = client.DownloadString(url);
            }

            return htmlCode;
        }
    }
}