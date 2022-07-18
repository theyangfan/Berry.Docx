using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    public class TableStyle : Style
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Style _style;
        private readonly TableRegionStyle _wholeTable;
        private readonly TableRegionStyle _firstRow;
        private readonly TableRegionStyle _lastRow;
        private readonly TableRegionStyle _firstColumn;
        private readonly TableRegionStyle _lastColumn;
        #endregion

        #region Constructors
        public TableStyle(Document doc, string styleName) : this(doc, StyleGenerator.GenerateTableStyle(doc))
        {
            base.Name = styleName;
            base.IsCustom = true;
            base.AddToGallery = true;
        }

        internal TableStyle(Document doc, W.Style style) : base(doc, style)
        {
            _doc = doc;
            _style = style;
            _wholeTable = new TableRegionStyle(doc, style, TableRegionType.WholeTable);
            _firstRow = new TableRegionStyle(doc, style, TableRegionType.FirstRow);
            _lastRow = new TableRegionStyle(doc, style, TableRegionType.LastRow);
            _firstColumn = new TableRegionStyle(doc, style, TableRegionType.FirstColumn);
            _lastColumn = new TableRegionStyle(doc, style, TableRegionType.LastColumn);
        }
        #endregion

        #region Public Properties
        public TableRegionStyle WholeTable => _wholeTable;

        public TableRegionStyle FirstRow => _firstRow;
        public TableRegionStyle LastRow => _lastRow;
        public TableRegionStyle FirstColumn => _firstColumn;

        public TableRegionStyle LastColumn => _lastColumn;

        #endregion

    }
}
