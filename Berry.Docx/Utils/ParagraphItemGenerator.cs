using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx
{
    internal class ParagraphItemGenerator
    {
        public static Break GenerateBreak()
        {
            Run run = new Run();
            Break br = new Break();
            run.AddChild(br);
            return br;
        }
    }
}
