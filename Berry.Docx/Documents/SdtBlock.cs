using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    public class SdtBlock : DocumentItem
    {
        private readonly SdtBlockContent _sdtContent;
        private readonly SdtBlockFormat _sdtProperties;

        public SdtBlock(Document doc): this(doc, SdtBlockGenerator.Generate(doc))
        {

        }

        internal SdtBlock(Document doc, W.SdtBlock sdt) : base(doc, sdt)
        {
            if (sdt.SdtContentBlock == null) sdt.SdtContentBlock = new W.SdtContentBlock();
            _sdtContent = new SdtBlockContent(doc, sdt.SdtContentBlock);
            if (sdt.SdtProperties == null) sdt.SdtProperties = new W.SdtProperties();
            _sdtProperties = new SdtBlockFormat(doc, sdt.SdtProperties);
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.SdtBlock;

        public SdtBlockFormat Format => _sdtProperties;

        public SdtBlockContent Content => _sdtContent;
    }
}
