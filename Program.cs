using System;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Text.Encodings.Web;
using System.Text.Unicode;
using System.Text;

namespace dotnet
{
    class Program
    {
        static void Main(string[] args)
        {
            parse();
            //parseExel();
        }

        // Download schedule and parse
        public static string parse()
        {
            WebClient client = new WebClient();
            using (Stream stream = client.OpenRead("http://www.mggeu.ru/wp-content/uploads/raspisanie/ochnoe/FPMiI_2021.xlsx"))
            {
                using (XLWorkbook wb = new XLWorkbook(stream))
                {
                    System.Console.WriteLine("Downloaded");
                    return parsewb(wb);
                }
            }
        }

        // Without downloading
        public static string parseExel()
        {
            const string fileName =  "ras.xlsx";
            var wb = new XLWorkbook(fileName);
            return parsewb(wb);
        }


        // Parse excel and return and save schedule as Json
        public static string parsewb(XLWorkbook wb)
        {
            string groupName = "ПИ-0319";
            var ws = wb.Worksheet(groupName);
            var wsran = ws.RangeUsed();
            int cols = wsran.ColumnCount();
            int rows = wsran.RowCount();
            
            Regex regexyear = new Regex(@"20\d\d");
            MatchCollection matches = regexyear.Matches(wsran.Cell(1, 1).GetString());
            int year = Convert.ToInt32(matches[0].Value);
            int month = 1;
            int day = 1;
            
            Rasp rasp = new Rasp(groupName);

            for (int i = 3; i <= cols; i+=2)
            {
                //Para para;
                int j = 2;
                while (j <= rows)
                {
                    
                    if((j-2) % 15 == 0)
                    {
                        if(!wsran.Cell(j, i).IsEmpty())
                        { 
                            var cell = wsran.Cell(j, i);
                            string date = cell.GetString();
                            string[] words = date.Split(new char[] {'.'}, StringSplitOptions.RemoveEmptyEntries);
                            month = Convert.ToInt32(words[1]);
                            day = Convert.ToInt32(words[0]);
                        }
                        j+=1;
                    }
                    else
                    {
                        if(!wsran.Cell(j, i).IsEmpty())
                        {
                            var cell = wsran.Cell(j, i);
                            string sbj = cell.GetString();
                            string teacher = cell.CellBelow().GetString();
                            string cab = cell.CellRight().GetString();

                            string[] words = wsran.Cell(j, 1).GetString()
                                .Split(new char[] {':'}, StringSplitOptions.RemoveEmptyEntries);

                            int hour = Convert.ToInt32(words[0]);
                            int minute = Convert.ToInt32(words[1].Substring(0, 2));
                            DateTime dtm = new DateTime (year, month, day, hour, minute, 0);
                            
                            //para = new Para(sbj, teacher, cab, dtm);
                            rasp.AddPara(dtm, sbj, teacher, cab);
                        }
                        j+=2;
                    }
                }
            }
            
            //rasp.PrintRasp();

            string jsonString = rasp.SerializeRasp();
            rasp.SaveAsJson();
            return jsonString;
        }
    }
}
