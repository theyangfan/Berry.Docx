using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Berry.Docx.Field
{
    public class Diagram : DrawingItem
    {
        internal Diagram(Document doc, W.Run ownerRun, W.Drawing drawing) : base(doc, ownerRun, drawing)
        {
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Diagram;
    }
}
