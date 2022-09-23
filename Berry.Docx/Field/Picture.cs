using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

using P = DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

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

        public Stream Stream
        {
            get
            {
                A.Blip blip = _drawing.Descendants<A.Blip>().FirstOrDefault();
                if (blip == null) return null;
                string rId = blip.Embed;
                P.ImagePart imagePart = (P.ImagePart)_doc.Package.MainDocumentPart.GetPartById(rId);
                return imagePart?.GetStream();
            }
        }
#endregion
    }
}
