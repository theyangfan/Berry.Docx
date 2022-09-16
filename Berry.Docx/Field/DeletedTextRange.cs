using System;
using System.Collections.Generic;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    public class DeletedTextRange : ParagraphItem
    {
        private readonly W.DeletedText _text;
        internal DeletedTextRange(Document doc, W.Run ownerRun) : base(doc, ownerRun, ownerRun.GetFirstChild<W.DeletedText>())
        {
            _text = ownerRun.GetFirstChild<W.DeletedText>();
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.DeletedTextRange;

        public string Text
        {
            get
            {
                return _text?.Text;
            }
        }
    }
}
