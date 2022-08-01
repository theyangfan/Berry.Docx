using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Berry.Docx.Field
{
    public class Picture : DrawingItem
    {
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.Drawing _drawing;
        private readonly W.Picture _picture;
        internal Picture(Document doc, W.Run ownerRun, W.Drawing drawing) : base(doc, ownerRun, drawing)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _drawing = drawing;
        }

        internal Picture(Document doc, W.Run ownerRun, W.Picture picture) : base(doc, ownerRun, picture)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _picture = picture;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Picture;

        #region Public Methods
        /// <summary>
        /// Creates a duplicate of the object.
        /// </summary>
        /// <returns>The cloned object.</returns>
        public override DocumentObject Clone()
        {
            W.Run run = new W.Run();
            run.RunProperties = _ownerRun.RunProperties?.CloneNode(true) as W.RunProperties; // copy format
            if(_drawing != null)
            {
                W.Drawing drawing = (W.Drawing)_drawing.CloneNode(true);
                run.AppendChild(drawing);
                return new Picture(_doc, run, drawing);
            }
            else
            {
                W.Picture pic = (W.Picture)_picture.CloneNode(true);
                run.AppendChild(pic);
                return new Picture(_doc, run, pic);
            }
        }
        #endregion
    }
}
