using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExercicioFinal_Jonatas
{
    class Connection
    {

        public static string SLToken;

        public static string Path;

        public static string CoockieName = "B1SESSION";

        public static string Url;

        public static string ServerName;

        public static void Context(string context, string url, string server)
        {

            Url = url;

            ServerName = server;

            string[] con = context.Split(';');

            foreach (string text in con)
            {
                if (text.Contains($"{CoockieName}="))
                {
                    SLToken = text.Replace($"{CoockieName}=", "");
                }

                if (text.Contains("path="))
                {
                    Path = text.Replace("path=", "");
                }
            }

        }
    }
}
