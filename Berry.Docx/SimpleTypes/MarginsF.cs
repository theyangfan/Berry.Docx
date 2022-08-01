using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    public class MarginsF
    {
        public MarginsF() { }
        public MarginsF(float left, float right, float top, float bottom)
        {
            Left = left;
            Right = right;
            Top = top;
            Bottom = bottom;
        }
        public float Left { get; set; }
        public float Right { get; set; }
        public float Top { get; set; }
        public float Bottom { get; set; }

        public override string ToString()
        {
            return "{" + $"Left:{Left}, Right:{Right}, Top:{Top}, Bottom:{Bottom}" + "}";
        }
    }
}
