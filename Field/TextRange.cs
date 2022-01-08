using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OO = DocumentFormat.OpenXml;
using OW = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;

namespace Berry.Docx.Field
{
    public class TextRange : DocumentObject
    {
        private OW.Run _run = null;
        public TextRange(Document doc, OW.Run run) : base(doc, run)
        {
            _run = run;
        }

        public string Text
        {
            get => _run.InnerText;
        }
    }
}
