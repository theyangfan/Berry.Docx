using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Berry.Docx.VisualModel
{
    internal static class ExtendedMethods
    {
        public static double ToPixel(this double point)
        {
            return point / 72 * 96;
        }

        public static float ToPixel(this float point)
        {
            return point / 72 * 96;
        }
    }
}
