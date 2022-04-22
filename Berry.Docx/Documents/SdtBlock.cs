using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    public class SdtBlock : DocumentItem
    {
        private readonly SdtContent _sdtContent;
        private readonly SdtProperties _sdtProperties;
        internal SdtBlock(Document doc, W.SdtBlock sdt) : base(doc, sdt)
        {
            if (sdt.SdtContentBlock != null)
                _sdtContent = new SdtContent(doc, sdt.SdtContentBlock);
            if (sdt.SdtProperties != null)
                _sdtProperties = new SdtProperties(doc, sdt.SdtProperties);
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.SdtBlock;

        public SdtContent SdtContent => _sdtContent;

        public SdtProperties SdtProperties => _sdtProperties;
    }
}
