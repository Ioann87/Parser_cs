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

namespace Test.Properties
{
    public class Comment
    {
        public string Name { set; get; }
        public string Date { set; get; }
        public string Text { set; get; }
        public int Raiting { set; get; }
        public int Like { set; get; }
        public int Dislike { set; get; }

        public static void WriteTo(JToken commentRes, List<Comment> comments)
        {
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
    }
}