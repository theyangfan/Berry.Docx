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

        /// <summary>
        /// Gets the cell borders.
        /// </summary>
        public TableBorders Borders => new TableBorders(_ownerDoc, this);

        /// <summary>
        /// Gets or sets the table cell vertical alignment.
        /// </summary>
        public TableCellVerticalAlignment VerticalCellAlignment
        {
            get
            {
                W.TableCellVerticalAlignment vAlign = _cell.TableCellProperties?.TableCellVerticalAlignment;
                if(vAlign == null)
                {
                    return vAlign.Val.Value.Convert<TableCellVerticalAlignment>();
                }
                return _ownerTable.GetStyle().WholeTable.VerticalCellAlignment;
            }
            set
            {
                if(_cell.TableCellProperties == null)
                {
                    _cell.TableCellProperties = new W.TableCellProperties();
                }
                _cell.TableCellProperties.TableCellVerticalAlignment = new W.TableCellVerticalAlignment()
                {
                    Val = value.Convert<W.TableVerticalAlignmentValues>()
                };
            }
        }

        public int GridSpan
        {
            get
            {
                W.GridSpan gridSpan = _cell.TableCellProperties?.GridSpan;
                if (gridSpan == null) return 1;
                return gridSpan.Val.Value;
            }
            set
            {
                if (value <= 1 && _cell.TableCellProperties?.GridSpan != null)
                {
                    _cell.TableCellProperties.GridSpan = null;
                    return;
                }
                if(_cell.TableCellProperties == null)
                {
                    _cell.TableCellProperties = new W.TableCellProperties();
                }
                if (_cell.TableCellProperties.GridSpan == null)
                {
                    _cell.TableCellProperties.GridSpan = new W.GridSpan(); 
                }
                _cell.TableCellProperties.GridSpan.Val = value;
            }
        }

        public TableCellVerticalMergeType VMerge
        {
            get
            {
                W.VerticalMerge vMerge = _cell.TableCellProperties?.VerticalMerge;
                if (vMerge == null) return TableCellVerticalMergeType.None;
                if (vMerge.Val == null) return TableCellVerticalMergeType.Continue;
                return vMerge.Val.Value.Convert<TableCellVerticalMergeType>();
            }
            set
            {
                if(value == TableCellVerticalMergeType.None && _cell.TableCellProperties?.VerticalMerge != null)
                {
                    _cell.TableCellProperties.VerticalMerge = null;
                    return;
                }
                if (_cell.TableCellProperties == null)
                {
                    _cell.TableCellProperties = new W.TableCellProperties();
                }
                if (_cell.TableCellProperties.VerticalMerge == null)
                {
                    _cell.TableCellProperties.VerticalMerge = new W.VerticalMerge();
                }
                if (value == TableCellVerticalMergeType.Restart)
                {
                    _cell.TableCellProperties.VerticalMerge.Val = W.MergedCellValues.Restart;
                }
                else if (value == TableCellVerticalMergeType.Continue)
                {
                    _cell.TableCellProperties.VerticalMerge.Val = null;
                }
            }
        }

        #endregion

        #region Public Methods
        /// <summary>
        /// Gets the width of the current cell.
        /// </summary>
        /// <returns></returns>
        public TableCellWidth GetCellWidth()
        {
            float width = 0;
            W.TableCellWidth tcWidth = _cell.TableCellProperties?.TableCellWidth;
            if (tcWidth?.Type == null) return new TableCellWidth(0, CellWidthType.Auto);
            float.TryParse(tcWidth.Width, out width);
            if(tcWidth.Type.Value == W.TableWidthUnitValues.Pct)
            {
                return new TableCellWidth(width / 50.0f, CellWidthType.Percent);
            }
            else if(tcWidth.Type.Value == W.TableWidthUnitValues.Dxa)
            {
                return new TableCellWidth(width / 20.0f, CellWidthType.Point);
            }
            return new TableCellWidth(0, CellWidthType.Auto);
        }

        /// <summary>
        /// Sets the width of the current cell.
        /// </summary>
        /// <param name="width">The cell width.</param>
        /// <param name="cellWidthType">The measurement type of the width.  </param>
        public void SetCellWidth(float width, CellWidthType cellWidthType)
        {
            if(_cell.TableCellProperties == null)
            {
                _cell.TableCellProperties = new W.TableCellProperties();
            }
            if(cellWidthType == CellWidthType.Auto)
            {
                _cell.TableCellProperties.TableCellWidth = new W.TableCellWidth() { Width = "0", Type = W.TableWidthUnitValues.Auto };
            }
            else if(cellWidthType == CellWidthType.Percent)
            {
                int percent = (int)Math.Round(width * 50);
                _cell.TableCellProperties.TableCellWidth = new W.TableCellWidth() { Width = percent.ToString(), Type = W.TableWidthUnitValues.Pct };
            }
            else
            {
                int w = (int)Math.Round(width * 20);
                _cell.TableCellProperties.TableCellWidth = new W.TableCellWidth() { Width = w.ToString(), Type = W.TableWidthUnitValues.Dxa };
            }
        }

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
            TableCell cell = (TableCell)Clone();
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
            TableCell cell = (TableCell)Clone();
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

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override DocumentObject Clone()
        {
            W.TableCell newCell = (W.TableCell)_cell.CloneNode(true);
            return new TableCell(_ownerDoc, _ownerTable, _ownerTableRow, newCell);
        }
        #endregion

        #region Internal
        internal new W.TableCell XElement => _cell;
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

    public class TableCellWidth
    {
        public TableCellWidth(float val, CellWidthType type)
        {
            Val = val;
            Type = type;
        }

        public float Val { get; set; }
        public CellWidthType Type { get; set; }
    }
}
