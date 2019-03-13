using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            using(ExcelPackage pkg = new ExcelPackage( new System.IO.FileInfo(@"test.xlsx")))
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
                for(int i = 0; i < 3; i++)
                {
                    dictionary.Add(i, list[i]);
                }

                ws.InsertTable(6, 2, dictionary, "Key-Value", ExcelColor.Danger, true);


                /// Insert hierarchical list
                var item01 = new HierarchyElement("Vízszintes átméretezés");
                var item02 = new HierarchyElement("Függőleges átméretezés");
                var item03 = new HierarchyElement("Átméretezés");
                item03.AddChild(item01);
                item03.AddChild(item02);
                var item04 = new HierarchyElement("Eltolás");
                var item05 = new HierarchyElement("Tranzformációk");
                item05.AddChild(item03);
                item05.AddChild(item04);
                var item06 = new HierarchyElement("O1");
                var item07 = new HierarchyElement("O2");
                var item08 = new HierarchyElement("Osztályozók");
                item08.AddChild(item06);
                item08.AddChild(item07);
                var item09 = new HierarchyElement("Aláíráshitelesítő");
                item09.AddChild(item05);
                item09.AddChild(item08);
                var root = new HierarchyElement("Benchmark");
                root.AddChild(item09);

                ws.InsertHierarchicalList(9, 2, root, "Hierarchical list", ExcelColor.Secondary);


                /// Insert legend 
                ws.Cells["B17:M23"].InsertLegend(
                    "Ez egy általános leírás arról, hogym it tud ez a táblázat  Ez egy általános leírás arról, hogym it tud ez a táblázat Ez egy általános leírás arról, hogym it tud ez a táblázat Ez egy általános leírás ",
                    "Bemutatkozó", true);
                

                /// Insert link
                if (pkg.Workbook.Worksheets["Sheet2"] == null)
                    pkg.Workbook.Worksheets.Add("Sheet2");

                ws.Cells["A1"].Value = "To Sheet2";
                ws.Cells["A1"].InsertLink("Sheet2");
                ws.Cells["B1"].Value = "To B2 in Sheet2";
                ws.Cells["B1"].InsertLink("Sheet2", "B2");

                pkg.Save();
                Console.WriteLine("Done");
                Console.ReadKey();
            }
        }
    }
}
