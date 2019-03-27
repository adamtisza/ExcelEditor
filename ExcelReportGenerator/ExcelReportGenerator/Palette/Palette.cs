using System.Drawing;

namespace ExcelReportGenerator
{

    class Palette
    {
        public Color MainColor { get; set; }
        public Color DarkColor { get; set; }
        public Color LightColor { get; set; }

        public Palette(Color main, Color dark, Color light)
        {
            MainColor = main;
            DarkColor = dark;
            LightColor = light;
        }

        public override string ToString()
        {
            return $"Palette: MainColor : {MainColor.ToString()}";
        }
    }




}
