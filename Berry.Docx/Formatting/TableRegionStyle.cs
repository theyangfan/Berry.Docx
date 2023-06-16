using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the style of one region in the table. Look <see cref="TableStyle"/> for supporting regions.
    /// <para>表示表格中某一区域的样式. <see cref="TableStyle"/> 中定义了支持的表格区域.</para>
    /// </summary>
    public class TableRegionStyle
    {
        #region Private Members
        private readonly Document _doc;
        private readonly Style _style;
        private readonly TableRegionType _region;
        private readonly TablePropertiesHolder _tblPr;
        private readonly TablePropertiesHolder _wholeTblPr;
        private readonly CharacterFormat _cFormat;
        private readonly ParagraphFormat _pFormat;
        private readonly TableBorders _borders;
        #endregion

        #region Constructors
        internal TableRegionStyle(Document doc, Style style, TableRegionType region)
        {
            _doc = doc;
            _style = style;
            _region = region;
            _tblPr = new TablePropertiesHolder(style, region);
            _wholeTblPr = new TablePropertiesHolder(style, TableRegionType.WholeTable);

            _cFormat = new CharacterFormat(doc, style.XElement, region);
            _pFormat = new ParagraphFormat(doc, style.XElement, region);
            _borders = new TableBorders(style, region);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the table cell character format.
        /// </summary>
        public CharacterFormat CharacterFormat => _cFormat;

        /// <summary>
        /// Gets the table cell paragraph format.
        /// </summary>
        public ParagraphFormat ParagraphFormat => _pFormat;

        /// <summary>
        /// Gets or sets the horizontal alignment.
        /// </summary>
        public TableRowAlignment HorizontalAlignment
        {
            get
            {
                if (_tblPr.HorizontalAlignment != null) 
                    return _tblPr.HorizontalAlignment;
                if (_region != TableRegionType.WholeTable && _wholeTblPr.HorizontalAlignment != null)
                    return _wholeTblPr.HorizontalAlignment;
                if (_style.BaseStyle != null) 
                    return new TableRegionStyle(_doc, _style.BaseStyle, _region).HorizontalAlignment;
                return TableRowAlignment.Left;
            }
            set
            {
                _tblPr.HorizontalAlignment = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether allow row to break across pages.
        /// </summary>
        public bool AllowBreakAcrossPages
        {
            get
            {
                if (_tblPr.AllowBreakAcrossPages != null) 
                    return _tblPr.AllowBreakAcrossPages;
                if (_region != TableRegionType.WholeTable && _wholeTblPr.AllowBreakAcrossPages != null)
                    return _wholeTblPr.AllowBreakAcrossPages;
                if (_style.BaseStyle != null)
                    return new TableRegionStyle(_doc, _style.BaseStyle, _region).AllowBreakAcrossPages;
                return true;
            }
            set
            {
                if (_region != TableRegionType.WholeTable) return;
                _tblPr.AllowBreakAcrossPages = value;
            }
        }

        /// <summary>
        /// Gets or sets the table cell vertical alignment.
        /// </summary>
        public TableCellVerticalAlignment VerticalCellAlignment
        {
            get
            {
                if (_tblPr.VerticalCellAlignment != null)
                    return _tblPr.VerticalCellAlignment;
                if (_region != TableRegionType.WholeTable && _wholeTblPr.VerticalCellAlignment != null)
                    return _wholeTblPr.VerticalCellAlignment;
                if (_style.BaseStyle != null)
                    return new TableRegionStyle(_doc, _style.BaseStyle, _region).VerticalCellAlignment;
                return TableCellVerticalAlignment.Top;
            }
            set
            {
                _tblPr.VerticalCellAlignment = value;
            }
        }

        /// <summary>
        /// Gets or sets the background color.
        /// </summary>
        public ColorValue Background
        {
            get
            {
                if (_tblPr.Background != null)
                    return _tblPr.Background;
                if (_region != TableRegionType.WholeTable && _wholeTblPr.Background != null)
                    return _wholeTblPr.Background;
                if (_style.BaseStyle != null)
                    return new TableRegionStyle(_doc, _style.BaseStyle, _region).Background;
                return ColorValue.Auto;
            }
            set
            {
                _tblPr.Background = value;
            }
        }

        /// <summary>
        /// Gets the table cell borders.
        /// </summary>
        public TableBorders Borders => _borders;
        #endregion
    }

    internal enum TableRegionType
    {
        WholeTable = 0,
        FirstRow = 1,
        LastRow = 2,
        FirstColumn = 3,
        LastColumn = 4
    }
}
