using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    public class Spacing
    {
        public Spacing() { }
        public Spacing(float val, SpacingUnit unit)
        {
            Val = val;
            Unit = unit;
        }
        public float Val { get; set; }

        public SpacingUnit Unit { get; set; }

        public override string ToString()
        {
            return $"Spacing[{Val} {Unit}]";
        }
    }
}
