using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text;

using P = DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represent a picture in the paragraph.
    /// </summary>
    public class Picture : DrawingItem
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.Drawing _drawing;
        private readonly W.Picture _picture;
        #endregion

        #region Constructors
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
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the type of the current object.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Picture;

        /// <summary>
        /// Gets the picture data stream.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the width of the picture.
        /// </summary>
        public int Width
        {
            get
            {
                if(_drawing != null)
                {
                    Wp.Extent extent = _drawing.Descendants<Wp.Extent>().FirstOrDefault();
                    if(extent != null)
                    {
                        return (int)Math.Round(extent.Cx.Value / 12700.0);
                    }
                }
                return 0;
            }
            set
            {
                if (_drawing != null)
                {
                    Wp.Extent extent = _drawing.Descendants<Wp.Extent>().FirstOrDefault();
                    if (extent != null)
                    {
                        extent.Cx = value * 12700;
                    }
                    A.Extents extents = _drawing.Descendants<A.Extents>().FirstOrDefault();
                    if(extents != null)
                    {
                        extents.Cx = value * 12700;
                    }
                }
            }
        }

        /// <summary>
        /// ets or sets the height of the picture.
        /// </summary>
        public int Height
        {
            get
            {
                if (_drawing != null)
                {
                    Wp.Extent extent = _drawing.Descendants<Wp.Extent>().FirstOrDefault();
                    if (extent != null)
                    {
                        return (int)Math.Round(extent.Cy.Value / 12700.0);
                    }
                }
                return 0;
            }
            set
            {
                if (_drawing != null)
                {
                    Wp.Extent extent = _drawing.Descendants<Wp.Extent>().FirstOrDefault();
                    if (extent != null)
                    {
                        extent.Cy = value * 12700;
                    }
                    A.Extents extents = _drawing.Descendants<A.Extents>().FirstOrDefault();
                    if (extents != null)
                    {
                        extents.Cy = value * 12700;
                    }
                }
            }
        }
        #endregion

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
