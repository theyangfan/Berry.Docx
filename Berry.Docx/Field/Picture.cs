using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Berry.Docx.Field
{
    public class Picture : DrawingItem
    {
        private readonly W.Drawing _drawing;
        internal Picture(Document doc, W.Run ownerRun, W.Drawing drawing) : base(doc, ownerRun, drawing)
        {
            _drawing = drawing;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Picture;


    }
}
