using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    public class LineSpacing
    {
        public LineSpacing() { }
        public LineSpacing(float val, LineSpacingRule rule)
        {
            Val = val;
            Rule = rule;
        }
        public float Val { get; set; }

        public LineSpacingRule Rule { get; set; }

        public override string ToString()
        {
            return $"LineSpacing[{Val} {Rule}]";
        }
    }
}
