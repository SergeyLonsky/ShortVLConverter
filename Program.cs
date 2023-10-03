using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ShortVLConverter
{
    class Program
    {
        static Excel.Application excel_app = new Excel.Application();
        static void Main(string[] args)
        {
            
            ParseVL();
            Console.WriteLine("Parsing is done");
            Console.ReadLine();            
        }

        static void ParseVL()
        {
            excel_app.DefaultFilePath = Directory.GetCurrentDirectory();
            List<(string, string, string)> retCsv = new List<(string, string, string)>();
            foreach (var file in Directory.GetFiles(@".\Input", "*.xlsx"))
            //foreach (var file in Directory.GetFiles(@"e:\Projects\ИнноТех\ConvertFromZulu\ShortVLConverter\bin\Debug\Input\", "*.xlsx"))
            {
                Console.WriteLine($"Start parsing {file}");
                Excel.Workbook wbTarget = excel_app.Workbooks.Open($"{Directory.GetCurrentDirectory()}\\{file}");
                Excel.Worksheet wsTarget = wbTarget.Sheets[1];
                //int rowsCount = 154;
                int index = 2;
                while(!string.IsNullOrEmpty(GetValue(wsTarget, 1, index)))
                //for (int index = 2; index < rowsCount; index++)
                {
                    var tmpuuid = GetValue(wsTarget, 3, index).ToLower();
                    Guid uuid;
                    if (Guid.TryParse(tmpuuid, out uuid))
                    {
                        //var uuid = GetValue(wsTarget, 3, index).ToLower();
                        var tmpCoords = GetValue(wsTarget, 8, index);
                        if (tmpCoords.Contains("LINESTRING"))
                        {
                            var coords = tmpCoords.Replace("LINESTRING", "").Replace("(", "").Replace(")", "");
                            var coordsSplit = !string.IsNullOrEmpty(coords) ? coords.Split(',') : null;
                            if (coordsSplit != null)
                            {
                                foreach (string coordlatlon in coordsSplit)
                                {
                                    var latlon = coordlatlon.Split(' ');
                                    if (!retCsv.Contains((uuid.ToString(), latlon[0], latlon[1])))
                                    {
                                        retCsv.Add((uuid.ToString(), latlon[0], latlon[1]));
                                    }
                                }
                            }
                        }
                    }
                    index++;
                }
                wbTarget.Close();
                Console.WriteLine($"Parsing {file} is done");
                using (var output = System.IO.File.CreateText($"{Directory.GetCurrentDirectory()}\\{file.Replace("Input", "Output")}_SK11.csv"))
                {
                    Console.WriteLine($"Start create .\\Output\\{file}_SK11.csv");
                    output.Write("UID;Latitude;Longitude;angle");
                    output.WriteLine();
                    foreach (var item in retCsv)
                    {
                        output.Write($"{item.Item1};{item.Item2};{item.Item3};0");
                        output.WriteLine();
                    }
                    Console.WriteLine($"Create .\\Output\\{file}_SK11.csv is done");
                }
            }            
        }

        private static string GetValue(Excel.Worksheet worksheet, int row, int col)
        {
            return worksheet.Cells[col, row].Value2 != null ? worksheet.Cells[col, row].Value2.ToString() : string.Empty;
        }


    }
}
