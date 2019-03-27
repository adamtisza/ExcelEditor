using System.Collections.Generic;
using OfficeOpenXml;

namespace testproject.TextLevel
{
    enum TextLevel
    {
        Title,
        Heading1,
        Heading2,
        Heading3,
        Normal
    }
    class Level
    {
        private readonly Dictionary<TextLevel, int> TextLevels;


        public Level(ExcelWorkbook wb)
        {
            //CreateNamedStyles(wb);
        }

        public int GetLevel(TextLevel level)
        {
            return TextLevels[level];
        }

        private static void CreateNamedStyles(ExcelWorkbook wb)
        {
            
            var titleStyle = wb.Styles.CreateNamedStyle("Title").Style;
            titleStyle.Font.Bold = true;
            titleStyle.Font.Size = 18;

            var h1Style = wb.Styles.CreateNamedStyle("Heading1").Style;
            h1Style.Font.Bold = true;
            h1Style.Font.Size = 15;

            var h2Style = wb.Styles.CreateNamedStyle("Heading2").Style;
            h2Style.Font.Bold = true;
            h2Style.Font.Size = 13;

            var h3Style = wb.Styles.CreateNamedStyle("Heading3").Style;
            h3Style.Font.Bold = true;
            h3Style.Font.Size = 11;
        }

    }
}
