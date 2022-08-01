using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Berry.Docx.Field
{

    public abstract class DrawingItem : ParagraphItem
    {
        private readonly W.Drawing _drawing;
        private readonly W.Picture _picture;
        internal DrawingItem(Document doc, W.Run ownerRun, W.Drawing drawing) : base(doc, ownerRun, drawing)
        {
            _drawing = drawing;
        }
        internal DrawingItem(Document doc, W.Run ownerRun, W.Picture picture) : base(doc, ownerRun, picture)
        {
            _picture = picture;
        }
        public TextWrappingStyle TextWrappingStyle
        {
            get
            {
                if(_drawing != null)
                    return (_drawing.FirstChild is Wp.Inline) ? TextWrappingStyle.Inline : TextWrappingStyle.Floating;
                return TextWrappingStyle.Inline;
            }
        }
    }
}
