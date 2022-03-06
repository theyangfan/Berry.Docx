using Berry.Docx.Collections;
using System;
using System.Collections.Generic;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// Represent the table row.
    /// </summary>
    public class TableRow : DocumentItem
    {
        #region Private Members
        private Document _ownerDoc;
        private Table _ownerTable;
        private W.TableRow _row;
        #endregion

        #region Constructors
        internal TableRow(Document ownerDoc, Table ownerTable, W.TableRow row) : base(ownerDoc, row)
        {
            _ownerDoc = ownerDoc;
            _ownerTable = ownerTable;
            _row = row;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The DocumentObject type.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.TableRow;

        /// <summary>
        /// The child DocumentObjects.
        /// </summary>
        public override DocumentObjectCollection ChildObjects => Cells;

        /// <summary>
        /// The table cells.
        /// </summary>
        public TableCellCollection Cells => new TableCellCollection(_row, GetTableCells());
        #endregion

        #region Public Methods
        /// <summary>
        /// Adds new table cell to the end of row.
        /// </summary>
        /// <returns></returns>
        public TableCell AddCell()
        {
            TableCell cell = Cells.Last().Clone();
            cell.ClearContent();
            Cells.Add(cell);
            return cell;
        }

        /// <summary>
        /// Inserts a new row above current row. 
        /// </summary>
        /// <returns></returns>
        public TableRow InsertRowAbove()
        {
            TableRow row = this.Clone();
            row.ClearContent();
            _row.InsertBeforeSelf(row.XElement);
            return row;
        }

        /// <summary>
        /// Inserts a new row below current row. 
        /// </summary>
        /// <returns></returns>
        public TableRow InsertRowBelow()
        {
            TableRow row = this.Clone();
            row.ClearContent();
            _row.InsertAfterSelf(row.XElement);
            return row;
        }

        /// <summary>
        /// Clears cells contents.
        /// </summary>
        public void ClearContent()
        {
            foreach (TableCell cell in Cells)
                cell.ClearContent();
        }
        #endregion

        #region Internal Methods
        internal TableRow Clone()
        {
            W.TableRow newRow = (W.TableRow)_row.CloneNode(true);
            return new TableRow(_ownerDoc, _ownerTable, newRow);
        }
        #endregion

        #region Private Methods
        private IEnumerable<TableCell> GetTableCells()
        {
            foreach (W.TableCell cell in _row.Elements<W.TableCell>())
            {
                yield return new TableCell(_ownerDoc, _ownerTable, this, cell);
            }
        }
        #endregion

    }
}
