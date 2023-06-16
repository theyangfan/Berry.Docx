using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represents the table format.
    /// </summary>
    public class TableFormat
    {
        #region Private Members
        private readonly Document _doc;
        private readonly Table _table;
        private readonly TablePropertiesHolder _tblPr;
        #endregion

        #region Constructors
        internal TableFormat(Document doc, Table table)
        {
            _doc = doc;
            _table = table;
            _tblPr = new TablePropertiesHolder(table);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Specifies that the first row format shall be applied to the table.
        /// </summary>
        public bool FirstRowEnabled
        {
            get => _tblPr.FirstRowEnabled ?? true;
            set => _tblPr.FirstRowEnabled = value;
        }

        /// <summary>
        /// Specifies that the last row format shall be applied to the table.
        /// </summary>
        public bool LastRowEnabled
        {
            get => _tblPr.LastRowEnabled ?? false;
            set => _tblPr.LastRowEnabled = value;
        }

        /// <summary>
        /// Specifies that the first column format shall be applied to the table.
        /// </summary>
        public bool FirstColumnEnabled
        {
            get => _tblPr.FirstColumnEnabled ?? true;
            set => _tblPr.FirstColumnEnabled = value;
        }

        /// <summary>
        /// Specifies that the last column format shall be applied to the table.
        /// </summary>
        public bool LastColumnEnabled
        {
            get => _tblPr.LastColumnEnabled ?? false;
            set => _tblPr.LastColumnEnabled = value;
        }

        /// <summary>
        /// Gets or sets the table horizontal alignment.
        /// </summary>
        public TableRowAlignment HorizontalAlignment
        {
            get
            {
                if(_tblPr.HorizontalAlignment != null) return _tblPr.HorizontalAlignment;
                return _table.GetStyle().WholeTable.HorizontalAlignment;
            }
            set
            {
                _tblPr.HorizontalAlignment = value;
                foreach(var row in _table.Rows)
                {
                    row.HorizontalAlignment = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the background color of the current table.
        /// </summary>
        public ColorValue Background
        {
            get
            {
                if(_tblPr.Background != null) return _tblPr.Background;
                return _table.GetStyle().WholeTable.Background;
            }
            set
            {
                _tblPr.Background = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the table is floating.
        /// </summary>
        public bool WrapTextAround
        {
            get => _tblPr.WrapTextAround ?? false;
            set => _tblPr.WrapTextAround = value;
        }

        /// <summary>
        /// Gets or sets a value indicating whether repeat the first row as header row at the top of each page.
        /// </summary>
        public bool RepeatHeaderRow
        {
            get => _table.Rows[0].RepeatHeaderRow;
            set => _table.Rows[0].RepeatHeaderRow = value;
        }

        /// <summary>
        /// Gets the table borders.
        /// </summary>
        public TableBorders Borders => new TableBorders(_table);
        #endregion
    }
}
