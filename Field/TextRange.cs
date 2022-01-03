using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OO = DocumentFormat.OpenXml;
using OW = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    public class TextRange : DocumentObject
    {
        private OW.Run _run = null;
        public TextRange(OW.Run run) : base(run)
        {
            _run = run;
        }

        public string Text
        {
            get => _run.InnerText;
        }
    }
}
