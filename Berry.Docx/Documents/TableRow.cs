using Berry.Docx.Collections;
using System;
using System.Collections.Generic;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Formatting;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// Represent the table row.
    /// </summary>
    public class TableRow : DocumentItem
    {
        #region Private Members
        private Document _ownerDoc;
        private W.TableRow _row;
        private readonly TablePropertiesHolder _tblPr;
        #endregion

        #region Constructors
        internal TableRow(Document ownerDoc, W.TableRow row) : base(ownerDoc, row)
        {
            _ownerDoc = ownerDoc;
            _row = row;
            _tblPr = new TablePropertiesHolder(this);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// 
        /// </summary>
        public Table Table
        {
            get
            {
                if(_row.Parent is W.Table)
                {
                    return new Table(_ownerDoc, (W.Table)_row.Parent);
                }
                return null;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public int RowIndex
        {
            get
            {
                if(Table == null) return 0;
                return Table.Rows.IndexOf(this);
            }
        }
        /// <summary>
        /// The DocumentObject type.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.TableRow;

        /// <summary>
        /// The child DocumentObjects.
        /// </summary>
        public override DocumentObjectCollection ChildObjects => Cells;

        /// <summary>
        /// Gets the table cell at the specified column.
        /// </summary>
        /// <param name="column">The zero-based index.</param>
        /// <returns>The table cell at the specified index.</returns>
        public TableCell this[int column] => Cells[column];

        /// <summary>
        /// The table cells.
        /// </summary>
        public TableCellCollection Cells => new TableCellCollection(_row, GetTableCells());

        /// <summary>
        /// Gets or sets the row height (int points).
        /// </summary>
        public float Height
        {
            get
            {
                W.TableRowHeight trHeight = _row.TableRowProperties?.GetFirstChild<W.TableRowHeight>();
                if(trHeight?.Val == null) return 0f;
                return trHeight.Val.Value / 20.0f;
            }
            set
            {
                if(value <= 0)
                {
                    if(_row.TableRowProperties?.GetFirstChild<W.TableRowHeight>() != null)
                    {
                        W.TableRowHeight trH = _row.TableRowProperties.GetFirstChild<W.TableRowHeight>();
                        if (trH.HeightType == null) _row.TableRowProperties.RemoveChild(trH);
                        else _row.TableRowProperties.GetFirstChild<W.TableRowHeight>().Val = null;
                    }
                    return;
                }
                if(_row.TableRowProperties == null)
                {
                    _row.TableRowProperties = new W.TableRowProperties();
                }
                if(_row.TableRowProperties.GetFirstChild<W.TableRowHeight>() == null)
                {
                    _row.TableRowProperties.AddChild(new W.TableRowHeight()); 
                }
                W.TableRowHeight trHeight = _row.TableRowProperties.GetFirstChild<W.TableRowHeight>();
                trHeight.Val = (uint)(value * 20);
            }
        }

        /// <summary>
        /// Gets or sets the height type.
        /// </summary>
        public TableRowHeightType HeightType
        {
            get
            {
                W.TableRowHeight trHeight = _row.TableRowProperties?.GetFirstChild<W.TableRowHeight>();
                if (trHeight?.HeightType == null) return TableRowHeightType.Auto;
                return trHeight.HeightType.Value.Convert<TableRowHeightType>();
            }
            set
            {
                if (value == TableRowHeightType.Auto)
                {
                    if (_row.TableRowProperties?.GetFirstChild<W.TableRowHeight>() != null)
                    {
                        W.TableRowHeight trH = _row.TableRowProperties.GetFirstChild<W.TableRowHeight>();
                        if (trH.Val == null) _row.TableRowProperties.RemoveChild(trH);
                        else _row.TableRowProperties.GetFirstChild<W.TableRowHeight>().HeightType = null;
                    }
                    return;
                }
                if (_row.TableRowProperties == null)
                {
                    _row.TableRowProperties = new W.TableRowProperties();
                }
                if (_row.TableRowProperties.GetFirstChild<W.TableRowHeight>() == null)
                {
                    _row.TableRowProperties.AddChild(new W.TableRowHeight());
                }
                W.TableRowHeight trHeight = _row.TableRowProperties.GetFirstChild<W.TableRowHeight>();
                trHeight.HeightType = value.Convert<W.HeightRuleValues>();
            }
        }

        /// <summary>
        /// Gets or sets the horizontal alignment.
        /// </summary>
        public TableRowAlignment HorizontalAlignment
        {
            get
            {
                if(_tblPr.HorizontalAlignment != null) return _tblPr.HorizontalAlignment;
                if(Table == null) return TableRowAlignment.Left;
                return Table.Format.HorizontalAlignment;
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
                if (_tblPr.AllowBreakAcrossPages != null) return _tblPr.AllowBreakAcrossPages;
                if(Table == null) return true;
                return Table.GetStyle().WholeTable.AllowBreakAcrossPages;
            }
            set
            {
                _tblPr.AllowBreakAcrossPages = value;
            }
        }
        #endregion

        #region Internal
        /// <summary>
        /// Gets or sets a value indicating whether repeat the first row as header row at the top of each page.
        /// </summary>
        internal bool RepeatHeaderRow
        {
            get => _tblPr.RepeatHeaderRow ?? false;
            set => _tblPr.RepeatHeaderRow = value;
        }

        internal new W.TableRow XElement => _row;
        #endregion

        #region Public Methods
        /// <summary>
        /// Adds new table cell to the end of row.
        /// </summary>
        /// <returns></returns>
        public TableCell AddCell()
        {
            TableCell cell = (TableCell)Cells.Last().Clone();
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
            TableRow row = (TableRow)Clone();
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
            TableRow row = (TableRow)Clone();
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

        public override DocumentObject Clone()
        {
            W.TableRow newRow = (W.TableRow)_row.CloneNode(true);
            return new TableRow(_ownerDoc, newRow);
        }
        #endregion

        #region Private Methods
        private IEnumerable<TableCell> GetTableCells()
        {
            foreach (W.TableCell cell in _row.Elements<W.TableCell>())
            {
                yield return new TableCell(_ownerDoc, cell);
            }
        }
        #endregion

    }
}
