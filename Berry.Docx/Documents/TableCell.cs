using Berry.Docx.Collections;
using System;
using System.Collections.Generic;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// Represent the table cell.
    /// </summary>
    public class TableCell : DocumentItem
    {
        #region Private Members
        private Document _ownerDoc;
        private Table _ownerTable;
        private TableRow _ownerTableRow;
        private W.TableCell _cell;
        #endregion

        #region Constructors
        internal TableCell(Document ownerDoc, Table ownerTable, TableRow ownerTableRow, W.TableCell cell) : base(ownerDoc, cell)
        {
            _ownerDoc = ownerDoc;
            _ownerTable = ownerTable;
            _ownerTableRow = ownerTableRow;
            _cell = cell;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The DocumentObject type.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.TableCell;

        /// <summary>
        /// The child DocumentObjects.
        /// </summary>
        public override DocumentObjectCollection ChildObjects => Paragraphs;

        /// <summary>
        /// The child paragraphs.
        /// </summary>
        public ParagraphCollection Paragraphs => new ParagraphCollection(_cell, GetParagraphs());

        #endregion

        #region Public Methods

        /// <summary>
        /// Inserts a new row above.
        /// </summary>
        /// <returns>The new table row.</returns>
        public TableRow InsertRowAbove()
        {
            return _ownerTableRow.InsertRowAbove();
        }

        /// <summary>
        /// Inserts a new row below.
        /// </summary>
        /// <returns>The new table row.</returns>
        public TableRow InsertRowBelow()
        {
            return _ownerTableRow.InsertRowBelow();
        }

        /// <summary>
        /// Inserts a new column on the left.
        /// </summary>
        public void InsertColumnLeft()
        {
            int index = _ownerTableRow.Cells.IndexOf(this);
            TableCell cell = Clone();
            cell.ClearContent();

            foreach(TableRow row in _ownerTable.Rows)
            {
                if(row.Cells.Count > index)
                {
                    row.Cells.InsertAt(cell.Clone(), index);
                }
                else
                {
                    row.Cells.Add(cell.Clone());
                }
            }
        }

        /// <summary>
        /// Inserts a new column on the right.
        /// </summary>
        public void InsertColumnRight()
        {
            int index = _ownerTableRow.Cells.IndexOf(this);
            TableCell cell = Clone();
            cell.ClearContent();

            foreach (TableRow row in _ownerTable.Rows)
            {
                if (row.Cells.Count > index)
                {
                    row.Cells.InsertAt(cell.Clone(), index+1);
                }
                else
                {
                    row.Cells.Add(cell.Clone());
                }
            }
        }

        /// <summary>
        /// Clears cell contents.
        /// </summary>
        public void ClearContent()
        {
            for (int i = Paragraphs.Count - 1; i > 0; i--)
            {
                Paragraphs.RemoveAt(i);
            }
            Paragraphs.First().Text = "";
        }
        #endregion

        #region Internal
        internal TableCell Clone()
        {
            W.TableCell newCell = (W.TableCell)_cell.CloneNode(true);
            return new TableCell(_ownerDoc, _ownerTable, _ownerTableRow, newCell);
        }
        #endregion

        #region Private Methods
        private IEnumerable<Paragraph> GetParagraphs()
        {
            foreach (W.Paragraph p in _cell.Elements<W.Paragraph>())
            {
                yield return new Paragraph(_ownerDoc, p);
            }
        }
        #endregion

    }
}
