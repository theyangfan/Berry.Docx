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
        private W.TableCell _cell;
        private TablePropertiesHolder _tblPr;
        #endregion

        #region Constructors
        internal TableCell(Document ownerDoc, W.TableCell cell) : base(ownerDoc, cell)
        {
            _ownerDoc = ownerDoc;
            _cell = cell;
            _tblPr = new TablePropertiesHolder(this);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The owner table.
        /// </summary>
        public Table Table => Row?.Table;

        /// <summary>
        /// The owner table row.
        /// </summary>
        public TableRow Row
        {
            get
            {
                if(_cell.Parent is W.TableRow)
                    return new TableRow(_ownerDoc, (W.TableRow)_cell.Parent);
                return null;
            }
        }

        /// <summary>
        /// The row index of the current cell.
        /// </summary>
        public int RowIndex => Row?.RowIndex ?? 0;

        /// <summary>
        /// The column index of the current cell.
        /// </summary>
        public int ColumnIndex
        {
            get
            {
                if (Row == null) return 0;
                int index = 0;
                foreach(var cell in Row.Cells)
                {
                    if (cell == this) break;
                    index += cell.ColumnSpan;
                }
                return index;
            }
        }

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
        public TableBorders Borders => new TableBorders(this);

        /// <summary>
        /// Gets or sets the cell background color.
        /// </summary>
        public ColorValue Background
        {
            get
            {
                if(_tblPr.Background != null)
                    return _tblPr.Background;
                if (Table == null) return ColorValue.Auto;
                if (RowIndex == 0 && Table.Format.FirstRowEnabled)
                {
                    var style = new TablePropertiesHolder(Table.GetStyle(), TableRegionType.FirstRow);
                    if(style.Background != null) return style.Background;
                }
                if (RowIndex == Table.RowCount - 1 && Table.Format.LastRowEnabled)
                {
                    var style = new TablePropertiesHolder(Table.GetStyle(), TableRegionType.LastRow);
                    if (style.Background != null) return style.Background;
                }
                if (ColumnIndex == 0 && Table.Format.FirstColumnEnabled)
                {
                    var style = new TablePropertiesHolder(Table.GetStyle(), TableRegionType.FirstColumn);
                    if (style.Background != null) return style.Background;
                }
                if (ColumnIndex + ColumnSpan == Table.ColumnCount && Table.Format.LastColumnEnabled)
                {
                    var style = new TablePropertiesHolder(Table.GetStyle(), TableRegionType.LastColumn);
                    if (style.Background != null) return style.Background;
                }
                return Table.Format.Background;
            }
            set
            {
                _tblPr.Background = value;
            }
        }

        /// <summary>
        /// Gets or sets the table cell vertical alignment.
        /// </summary>
        public TableCellVerticalAlignment VerticalCellAlignment
        {
            get
            {
                if(_tblPr.VerticalCellAlignment != null)
                    return _tblPr.VerticalCellAlignment;
                if (Table == null) return TableCellVerticalAlignment.Top;
                if (RowIndex == 0 && Table.Format.FirstRowEnabled)
                {
                    var style = new TablePropertiesHolder(Table.GetStyle(), TableRegionType.FirstRow);
                    if (style.VerticalCellAlignment != null) return style.VerticalCellAlignment;
                }
                if (RowIndex == Table.RowCount - 1 && Table.Format.LastRowEnabled)
                {
                    var style = new TablePropertiesHolder(Table.GetStyle(), TableRegionType.LastRow);
                    if (style.VerticalCellAlignment != null) return style.VerticalCellAlignment;
                }
                if (ColumnIndex == 0 && Table.Format.FirstColumnEnabled)
                {
                    var style = new TablePropertiesHolder(Table.GetStyle(), TableRegionType.FirstColumn);
                    if (style.VerticalCellAlignment != null) return style.VerticalCellAlignment;
                }
                if (ColumnIndex + ColumnSpan == Table.ColumnCount && Table.Format.LastColumnEnabled)
                {
                    var style = new TablePropertiesHolder(Table.GetStyle(), TableRegionType.LastColumn);
                    if (style.VerticalCellAlignment != null) return style.VerticalCellAlignment;
                }
                return Table.GetStyle().WholeTable.VerticalCellAlignment;
            }
            set
            {
                _tblPr.VerticalCellAlignment = value;
            }
        }

        /// <summary>
        /// Gets or sets the number of grid columns which shall be spanned by the current cell.
        /// </summary>
        public int ColumnSpan
        {
            get => _tblPr.ColumnSpan ?? 1;
            set => _tblPr.ColumnSpan = value;
        }

        /// <summary>
        /// Gets or sets a value indicates whether the current cell is part of a vertically 
        /// merged set of cells (i.e., whether this cell continues the vertical merge or starts
        /// a new merged group of cells).
        /// </summary>
        public TableCellVerticalMergeType VerticalMerge
        {
            get => _tblPr.VMerge ?? TableCellVerticalMergeType.None;
            set => _tblPr.VMerge = value;
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
            if (Row == null) return null;
            return Row.InsertRowAbove();
        }

        /// <summary>
        /// Inserts a new row below.
        /// </summary>
        /// <returns>The new table row.</returns>
        public TableRow InsertRowBelow()
        {
            if (Row == null) return null;
            return Row.InsertRowBelow();
        }

        /// <summary>
        /// Inserts a new column on the left.
        /// </summary>
        public void InsertColumnLeft()
        {
            if (Row == null) return;
            int index = Row.Cells.IndexOf(this);
            TableCell cell = (TableCell)Clone();
            cell.ClearContent();

            foreach(TableRow row in Table.Rows)
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
            if (Row == null) return;
            int index = Row.Cells.IndexOf(this);
            TableCell cell = (TableCell)Clone();
            cell.ClearContent();

            foreach (TableRow row in Table.Rows)
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
            return new TableCell(_ownerDoc, newCell);
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
