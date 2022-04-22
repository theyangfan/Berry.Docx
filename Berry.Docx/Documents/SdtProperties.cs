using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    public class SdtProperties : DocumentItem
    {
        internal SdtProperties(Document doc, W.SdtProperties sdtPr) : base(doc, sdtPr)
        {

        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.SdtProperties;
    }
}
