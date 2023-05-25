using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Berry.Docx.Visual
{
    public class Margin
    {
        public Margin(double left, double top, double right, double bottom)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        public double Left { get; set; } = 0;
        public double Top { get; set; } = 0;
        public double Right { get; set; } = 0;
        public double Bottom { get; set; } = 0;
    }
}
