using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGenerator
{
    class TestObject
    {
        public string Name { get; set; }
        public string Category { get; set; }
        public int Amount { get; set; }

        public TestObject(string Name, string Category, int Amount)
        {
            this.Name = Name;
            this.Category = Category;
            this.Amount = Amount;
        }

        public override string ToString()
        {
            return $"{Category}: {Name}, {Amount}db";
        }
    }


}
