using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    public class SpecialIndentation
    {
        public SpecialIndentation() { }
        public SpecialIndentation(SpecialIndentationType type, float val, IndentationUnit unit)
        {
            Val = val;
            Unit = unit;
            Type = type;
        }
        public float Val { get; set; }

        public IndentationUnit Unit { get; set; }

        public SpecialIndentationType Type { get; set; }

        public override string ToString()
        {
            if(Type != SpecialIndentationType.None)
                return $"SpecialIndentation[{Type}: {Val} {Unit}]";
            else
                return $"SpecialIndentation[{Type}]";
        }
    }
}
