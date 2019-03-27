using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace ExcelReportGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelPackage pkg = new ExcelPackage(new System.IO.FileInfo(@"test.xlsx")))
            {

                ExcelWorksheet ws;

                if (pkg.Workbook.Worksheets["Sheet1"] == null)
                    ws = pkg.Workbook.Worksheets.Add("Sheet1");
                else
                    ws = pkg.Workbook.Worksheets["Sheet1"];

                ///Clear worksheet for testing
                ws.Cells["A1:Z100"].Clear();


                /// Insert table with 2D array
                string[,] data = {
                { "ID", "Name", "Color"},
                { "A1", "Chocolate", "Brown"},
                { "A2", "Milk", "White"},
                };


                ws.InsertTable(2, 2, data, "2D Array");

                /// Insert table with IEnumerable
                var list = new List<TestObject>
                {
                    {new TestObject("Társasjáték", "Játék", 8)},
                    {new TestObject("Videojáték", "Játék", 10)},
                    {new TestObject("Kenyér", "Pékáru", 5)},
                    {new TestObject("Alma", "Gyümölcs", 10)},
                    {new TestObject("Autó", "Jármű", 1)},
                    {new TestObject("Éjjeliszekrény", "Bútor", 3)}
                };

                ws.InsertTable(2, 7, list, "IEnumerable", ExcelColor.Succes);

                /// Insert table with Key-Value pairs
                var dictionary = new Dictionary<int, TestObject>();
                for (int i = 0; i < 3; i++)
                {
                    dictionary.Add(i, list[i]);
                }

                ws.InsertTable(6, 2, dictionary, "Key-Value", ExcelColor.Danger);


                /// Insert hierarchical list

                var root = new HierarchyElement("Benchmark") {
                    new HierarchyElement("Aláíráshitelesítő")
                    {
                        new HierarchyElement("Transzformációk")
                        {
                            new HierarchyElement("Átmáretezés")
                            {
                                new HierarchyElement("Vízszintes átméretezés"),
                                new HierarchyElement("Függőleges átméretezés")
                            },
                            new HierarchyElement("Eltolás")
                        },
                        new HierarchyElement("Osztályozók")
                        {
                            new HierarchyElement("O1"),
                            new HierarchyElement("O2")
                        }
                    }
                };

                ws.InsertHierarchicalList(9, 2, root, "Hierarchical list", ExcelColor.Secondary);

                /// Insert legend 
                ws.Cells["B17:M23"].InsertLegend(
                    "Ez egy általános leírás arról, hogym it tud ez a táblázat  Ez egy általános leírás arról, hogym it tud ez a táblázat Ez egy általános leírás arról, hogym it tud ez a táblázat Ez egy általános leírás ",
                    "Bemutatkozó");


                /// Insert link
                if (pkg.Workbook.Worksheets["Sheet2"] == null)
                    pkg.Workbook.Worksheets.Add("Sheet2");

                ws.Cells["A1"].Value = "To Sheet2";
                ws.Cells["A1"].InsertLink("Sheet2");
                ws.Cells["B1"].Value = "To B2 in Sheet2";
                ws.Cells["B1"].InsertLink("Sheet2", "B2");

                ///Insert graphs

                var ws2 = pkg.Workbook.Worksheets["Sheet2"];
                ws2.Drawings.Clear();
                ws2.Cells["A1:Z100"].Clear();

                string[,] chartHeader = {
                    { "xLabel", "FAR", "FRR" }
                };

                double [,] chartData = {
                    {0.5, 0, 100 },
                    {0.9, 10, 80 },
                    {1.4, 30, 60 },
                    {1.9, 50, 50 },
                    {2.2, 70, 30 },
                    {2.7, 90, 20 },
                    {3, 100, 0 }
                };

                ws2.InsertTable(2, 2, chartHeader, null, ExcelColor.Transparent, false, false);
                ws2.InsertTable(2, 3, chartData, null, ExcelColor.Transparent, false, false);

                ws2.InsertLineChart(ws2.Cells["B3:D9"], 6, 1, "Error Rates", ws2.Cells["B2"].Value?.ToString(), "yLabel", ws2.Cells["C2:D2"], 500, 400);
                ws2.InsertColumnChart(ws2.Cells["B3:D9"], 14, 1, "Error Rates 2", ws2.Cells["B2"].Value?.ToString(), "yLabel", ws2.Cells["C2:D2"], 500, 400, "Error Rates");
                ws2.InsertColumnChart(ws2.Cells["B3:C9"], 6, 23, "Error Rates 3", ws2.Cells["B2"].Value?.ToString(), "yLabel", ws2.Cells["C2:D2"], 500, 400, "Error Rates");


                pkg.Save();
                Console.WriteLine("Done");
                Console.ReadKey();
            }
        }
    }
}
