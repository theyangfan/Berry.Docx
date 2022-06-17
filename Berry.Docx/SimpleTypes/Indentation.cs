using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    public class Indentation
    {
        public Indentation() { }
        public Indentation(float val, IndentationUnit unit)
        {
            Val = val;
            Unit = unit;
        }
        public float Val { get; set; }

        public IndentationUnit Unit { get; set; }

        public override string ToString()
        {
            return $"Indentation[{Val} {Unit}]";
        }
    }
}
