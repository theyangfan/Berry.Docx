﻿using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Berry.Docx.Field
{
    public class Shape : DrawingItem
    {
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.Drawing _drawing;

        internal Shape(Document doc, W.Run ownerRun, W.Drawing drawing) : base(doc, ownerRun, drawing)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _drawing = drawing;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Shape;

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
