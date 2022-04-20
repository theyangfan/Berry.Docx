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
        internal DrawingItem(Document doc, W.Run ownerRun, W.Drawing drawing) : base(doc, ownerRun, drawing)
        {
            _drawing = drawing;
        }
        public TextWrappingStyle TextWrappingStyle
        {
            get
            {
                return (_drawing.FirstChild is Wp.Inline) ? TextWrappingStyle.Inline : TextWrappingStyle.Floating;
            }
        }
    }
}
