using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

namespace Berry.Docx
{
    internal class ColorConverter
    {
        public static string ToHex(Color color)
        {
            string r = Convert.ToString(color.R, 16).PadLeft(2, '0').ToUpper();
            string g = Convert.ToString(color.G, 16).PadLeft(2, '0').ToUpper();
            string b = Convert.ToString(color.B, 16).PadLeft(2, '0').ToUpper();
            return $"{r}{g}{b}";
        }

        public static Color FromHex(string hex)
        {
            int r = 0, g = 0, b = 0;
            if (hex.Length == 6)
            {
                r = Convert.ToInt32(hex.Substring(0, 2), 16);
                g = Convert.ToInt32(hex.Substring(2, 2), 16);
                b = Convert.ToInt32(hex.Substring(4, 2), 16);
            }
            return Color.FromArgb(r, g, b);
        }
    }
}
