using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    public class Picture : DocumentItem
    {
        internal Picture(Document doc, W.Drawing drawing) : base(doc, drawing)
        {

        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Picture;
    }
}
