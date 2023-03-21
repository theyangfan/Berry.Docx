using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represent a shape in the document.
    /// </summary>
    public class Shape : DrawingItem
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.Drawing _drawing;
        #endregion

        #region Constructors
        internal Shape(Document doc, W.Run ownerRun, W.Drawing drawing) : base(doc, ownerRun, drawing)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _drawing = drawing;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the type of the current object.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Shape;
        #endregion

        #region Public Methods
        /// <summary>
        /// Creates a duplicate of the object.
        /// </summary>
        /// <returns>The cloned object.</returns>
        public override DocumentObject Clone()
        {
            W.Run run = new W.Run();
            W.Drawing drawing = (W.Drawing)_drawing.CloneNode(true);
            run.RunProperties = _ownerRun.RunProperties?.CloneNode(true) as W.RunProperties; // copy format
            run.AppendChild(drawing);
            return new Shape(_doc, run, drawing);
        }
        #endregion
    }
}
