using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using BF = Berry.Docx.Field;

namespace Berry.Docx.Visual.Field
{
    public class Picture : ParagraphItem
    {
        #region Private Members
        private Stream _stream;
        private double _width;
        private double _height;
        private HorizontalAlignment _hAlign = HorizontalAlignment.Left;
        #endregion

        #region Constructors
        internal Picture(BF.Picture picture)
        {
            _width = ((double)picture.Width).ToPixel();
            _height = ((double)picture.Height).ToPixel();
            _stream = picture.Stream;
        }
        #endregion

        #region Public Properties
        public Stream Stream => _stream;
        public override double Width => _width;
        public override double Height => _height;
        public override HorizontalAlignment HorizontalAlignment
        {
            get => _hAlign;
            internal set => _hAlign = value;
        }
        #endregion
    }
}
