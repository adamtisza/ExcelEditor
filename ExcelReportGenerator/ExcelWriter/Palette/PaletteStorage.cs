﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGenerator
{

    enum ExcelColor
    {
        Primary = 1,
        Secondary = 2,
        Succes = 3,
        Danger = 4,
        Warning = 5,
        Info = 6
    }

    class PaletteStorage
    {
        //Load colors into a dictionary
        private static readonly Dictionary<ExcelColor, Palette> PaletteColors = new Dictionary<ExcelColor, Palette>
        {
            {ExcelColor.Primary, new Palette(ColorFromRgb(91, 155, 213), ColorFromRgb(47, 117, 181), ColorFromRgb(155, 194, 230)) },
            {ExcelColor.Secondary, new Palette(ColorFromRgb(165, 165, 165), ColorFromRgb(123, 123, 123), ColorFromRgb(206, 206, 206)) },
            {ExcelColor.Succes, new Palette(ColorFromRgb(112, 173, 71), ColorFromRgb(84, 130, 53), ColorFromRgb(198, 224, 180)) },
            {ExcelColor.Danger, new Palette(ColorFromRgb(255, 59, 59), ColorFromRgb(208, 0, 0), ColorFromRgb(255, 174, 174)) },
            {ExcelColor.Warning, new Palette(ColorFromRgb(255, 217, 102), ColorFromRgb(255, 192, 0), ColorFromRgb(255, 230, 153)) },
            {ExcelColor.Info, new Palette(ColorFromRgb(92, 214, 234), ColorFromRgb(23, 162, 184), ColorFromRgb(170, 233, 244)) },
        };


        public static Palette GetPalette(ExcelColor excelColor)
        {
            return PaletteColors[excelColor];
        }


        //All used colors have an alpha of 255, no need to always call it explicitly
        private static Color ColorFromRgb(byte r, byte g, byte b)
        {
            return Color.FromArgb(255, r, g, b);
        }

    }
}
