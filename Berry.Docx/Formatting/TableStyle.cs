using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the table style. Supports get/set the wholeTable, first/last row/column region format.
    /// <para>表示一个段落样式，支持读写其整个表格、首行、末行、首列、末列区域的格式。</para>
    /// </summary>
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
        /// <summary>
        /// Creates a new table style with the specified name.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="styleName">The specified style name.</param>
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
        /// <summary>
        /// Gets the whole table format.
        /// </summary>
        public TableRegionStyle WholeTable => _wholeTable;

        /// <summary>
        /// Gets the first row format of the current table.
        /// </summary>
        public TableRegionStyle FirstRow => _firstRow;

        /// <summary>
        /// Gets the last row format of the current table.
        /// </summary>
        public TableRegionStyle LastRow => _lastRow;

        /// <summary>
        /// Gets the first column format of the current table.
        /// </summary>
        public TableRegionStyle FirstColumn => _firstColumn;

        /// <summary>
        /// Gets the last column format of the current table.
        /// </summary>
        public TableRegionStyle LastColumn => _lastColumn;
        #endregion

        #region Public Methods
        /// <summary>
        /// Gets the default table style.
        /// <para>获取默认表格样式.</para>
        /// </summary>
        /// <param name="doc"></param>
        /// <returns>The default table style.</returns>
        public static TableStyle Default(Document doc)
        {
            return doc.Styles.Where(s => s.Type == StyleType.Table && s.IsDefault).FirstOrDefault() as TableStyle;
        }
        #endregion
    }
}
