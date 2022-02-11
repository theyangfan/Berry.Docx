using System;
using System.Collections.Generic;
using System.Text;

using DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Utils
{
    internal class RunGenerator
    {
        public static Run Generate(string text)
        {
            Run run = new Run();

            RunProperties rPr = new RunProperties();
            RunFonts rFonts = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            rPr.AddChild(rFonts);

            Text text1 = new Text();
            text1.Text = text;

            run.AddChild(rPr);
            run.AddChild(text1);

            return run;
        }
    }
}
