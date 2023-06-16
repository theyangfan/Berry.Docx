using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace Berry.Docx.Visual
{
    public class Borders
    {
        public Border Top { get; set; } = new Border();
        public Border Bottom { get; set; } = new Border();
        public Border Left { get; set; } = new Border();
        public Border Right { get; set; } = new Border();
    }

    public class Border
    {
        public bool Visible { get; set; }

        public double Width { get; set; }

        public Color Color { get; set; } = Color.Black;
    }
}
