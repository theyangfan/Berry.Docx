using System.Collections.Generic;
using O = DocumentFormat.OpenXml;
using Berry.Docx.Documents;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// Represent a table cell collection.
    /// </summary>
    public class TableCellCollection : DocumentItemCollection, IEnumerable<TableCell>
    {
        #region Private Members
        private IEnumerable<TableCell> _cells;
        #endregion

        #region Constructors
        internal TableCellCollection(O.OpenXmlElement owner, IEnumerable<TableCell> cells)
            : base(owner, cells)
        {
            _cells = cells;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the table cell at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        /// <returns>The table cell at the specified index.</returns>
        public new TableCell this[int index] => (TableCell)base[index];
        #endregion

        #region Public Methods
        /// <summary>
        /// Returns the first table cell of the current collection.
        /// </summary>
        /// <returns>The first table cell in the current collection.</returns>
        public new TableCell First()
        {
            return (TableCell)base.First();
        }

        /// <summary>
        /// Returns the last table cell of the current collection.
        /// </summary>
        /// <returns>The last table cell in the current collection.</returns>
        public new TableCell Last()
        {
            return (TableCell)base.Last();
        }

        public IEnumerator<TableCell> GetEnumerator()
        {
            return _cells.GetEnumerator();
        }
        #endregion
    }
}
