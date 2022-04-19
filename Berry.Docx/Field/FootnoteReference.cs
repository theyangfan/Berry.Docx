using System;
using System.Collections.Generic;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    public class FootnoteReference : ParagraphItem
    {
        private readonly W.FootnoteReference _fnRef;
        internal FootnoteReference(Document doc, W.Run ownerRun, W.FootnoteReference fnRef)
            :base(doc, ownerRun, fnRef)
        {
            _fnRef = fnRef;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.FootnoteReference;

        public int Id
        {
            get
            {
                if (_fnRef.Id != null) return (int)_fnRef.Id;
                return -1;
            }
        }
    }
}
