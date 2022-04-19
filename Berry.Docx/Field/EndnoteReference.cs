using System;
using System.Collections.Generic;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    public class EndnoteReference : ParagraphItem
    {
        private readonly W.EndnoteReference _enRef;
        internal EndnoteReference(Document doc, W.Run ownerRun, W.EndnoteReference enRef)
            :base(doc, ownerRun, enRef)
        {
            _enRef = enRef;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.EndnoteReference;

        public int Id
        {
            get
            {
                if (_enRef.Id != null) return (int)_enRef.Id;
                return -1;
            }
        }
    }
}
