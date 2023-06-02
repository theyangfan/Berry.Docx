using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx.Visual
{
    public abstract class ParagraphItem
    {
        public abstract double Width { get; }
        public abstract double Height { get; }

        public abstract HorizontalAlignment HorizontalAlignment { get; internal set; }
    }
}
