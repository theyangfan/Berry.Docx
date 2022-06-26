using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    public class Borders
    {
        private readonly W.Run _run;
        private readonly W.Paragraph _paragraph;
        private readonly W.Style _style;
        private readonly BorderType _borderType;
        internal Borders(W.Run run)
        {
            _run = run;
        }

        internal Borders(W.Paragraph paragraph, BorderType type)
        {
            _paragraph = paragraph;
            _borderType = type;
        }

        internal Borders(W.Style style, BorderType type)
        {
            _style = style;
            _borderType = type;
        }
    }

    
}
