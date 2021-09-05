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
    public class Connection
    {
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