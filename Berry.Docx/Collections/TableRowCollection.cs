using System.Collections.Generic;
using System.Linq;
using O = DocumentFormat.OpenXml;
using Berry.Docx.Documents;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// Represent table row collection.
    /// </summary>
    public class TableRowCollection : DocumentItemCollection
    {
        #region Private Members
        private IEnumerable<TableRow> _rows;
        #endregion

        #region Constructors
        internal TableRowCollection(O.OpenXmlElement owner, IEnumerable<TableRow> rows)
            : base(owner, rows)
        {
            _rows = rows;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the table row at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        /// <returns>The table row at the specified index.</returns>
        public new TableRow this[int index] => (TableRow)base[index];
        #endregion

        #region Public Methods
        /// <summary>
        /// Returns the first table row of the current collection.
        /// </summary>
        /// <returns>The first table row in the current collection.</returns>
        public new TableRow First()
        {
            return (TableRow)base.First();
        }

        /// <summary>
        /// Returns the last table row of the current collection.
        /// </summary>
        /// <returns>The last table row in the current collection.</returns>
        public new TableRow Last()
        {
            return (TableRow)base.Last();
        }

        public IEnumerator<TableRow> GetEnumerator()
        {
            return _rows.GetEnumerator();
        }
        #endregion
    }
}
