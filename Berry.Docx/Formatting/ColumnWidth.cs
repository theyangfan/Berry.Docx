using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    public class ColumnWidth
    {
        #region Private Members
        private readonly W.GridColumn _column;
        #endregion

        #region Constructors
        internal ColumnWidth(Document doc, W.GridColumn column)
        {
            _column = column;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets or sets the width of the current column.
        /// </summary>
        public float Width
        {
            get
            {
                if (_column.Width == null) return 0;
                int.TryParse(_column.Width, out int width);
                return width / 20f;
            }
            set
            {
                _column.Width = Convert.ToInt32(value * 20).ToString();
            }
        }
        #endregion

    }
}
