using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Collections;
using Berry.Docx.Field;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// Represent the table.
    /// </summary>
    public class Table : DocumentItem
    {
        #region Private Members
        private Document _doc;
        private W.Table _table;
        #endregion

        #region Constructors
        /// <summary>
        /// The table constructor.
        /// </summary>
        /// <param name="doc">The owner document.</param>
        /// <param name="rowCnt">Table row count.</param>
        /// <param name="columnCnt">Table column count.</param>
        public Table(Document doc, int rowCnt, int columnCnt)
            : this(doc, TableGenerator.GenerateTable(rowCnt, columnCnt))
        {
        }

        internal Table(Document doc, W.Table table) : base(doc, table)
        {
            _doc = doc;
            _table = table;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The DocumentObject type.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Table;

        /// <summary>
        /// The child DocumentObjects of this table.
        /// </summary>
        public override DocumentObjectCollection ChildObjects => Rows;

        /// <summary>
        /// The table rows collection.
        /// </summary>
        public TableRowCollection Rows => new TableRowCollection(_table, TableRowsPrivate());
        #endregion

        #region Public Methods
        /// <summary>
        /// Adds a new row to the end of table.
        /// </summary>
        /// <returns>The table row.</returns>
        public TableRow AddRow()
        {
            TableRow row = Rows.Last().Clone();
            row.ClearContent();
            Rows.Add(row);
            return row;
        }
        #endregion

        #region Private Methods
        private IEnumerable<TableRow> TableRowsPrivate()
        {
            foreach (W.TableRow row in _table.Elements<W.TableRow>())
            {
                yield return new TableRow(_doc, this, row);
            }
        }
        #endregion
    }
}
