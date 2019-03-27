using System.ComponentModel.DataAnnotations;

namespace ExcelReportGenerator
{
    class TestObject
    {
        [Display(Name = "Név", Description = "A termék neve")]
        public string Name { get; set; }
        [Display(Name = "Kategória", Description ="A termék kategóriája")]
        public string Category { get; set; }
        [Display(Name = "Mennyiség", Description = "A rendelkezésre álló mennyiség")]
        [DisplayFormat(DataFormatString = "0.00")]
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
