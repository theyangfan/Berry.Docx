using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

namespace Berry.Docx
{
    internal class ColorConverter
    {
        /// <summary>
        /// Converts a <see cref="Color"/> value the the RRGGBB hex string.
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        public static string ToHex(Color color)
        {
            string r = Convert.ToString(color.R, 16).PadLeft(2, '0').ToUpper();
            string g = Convert.ToString(color.G, 16).PadLeft(2, '0').ToUpper();
            string b = Convert.ToString(color.B, 16).PadLeft(2, '0').ToUpper();
            return $"{r}{g}{b}";
        }

        /// <summary>
        /// Convert a RRGGBB hex string to the <see cref="Color"/>.
        /// </summary>
        /// <param name="hex">The hex string like RRGGBB.</param>
        /// <returns></returns>
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
