using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    public class Picture : ParagraphItem
    {
        private readonly W.Drawing _drawing;
        internal Picture(Document doc, W.Run ownerRun, W.Drawing drawing) : base(doc, ownerRun, drawing)
        {
            _drawing = drawing;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Picture;


    }
}
