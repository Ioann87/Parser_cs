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
            
            List<Comment> comments = new List<Comment>();

            while (check)
            {
                try
                {
                    var response = JObject.Parse(Connection.GetResponse(url));
                    var commentRes = response["data"]?["feedback"];

                    if ((int) commentRes["state"] == -1)
                    {
                        check = false;
                        break;
                    }

                    for (int i = 0; i < commentRes.Count(); i++)
                    {
                        Comment comment = new Comment();
                        comment.Name = (string) commentRes[i]["userName"];
                        comment.Date = (string) commentRes[i]["date"];
                        comment.Text = (string) commentRes[i]["text"];
                        comment.Raiting = (int) commentRes[i]["mark"];
                        comment.Like = (int) commentRes[i]["votesUp"];
                        comment.Dislike = (int) commentRes[i]["votesDown"];

                        comments.Add(comment);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                page++;
            }

            await ExcelWriter.InitTable(comments);
        }
    }
}