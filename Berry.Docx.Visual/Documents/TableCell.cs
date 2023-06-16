using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using B = Berry.Docx;
using BD = Berry.Docx.Documents;

namespace Berry.Docx.Visual.Documents
{
    public class TableCell
    {
        #region Private Members
        private int _rowIndex = 0;
        private int _colIndex = 0;
        private int _rowSpan = 1;
        private int _colSpan = 1;
        private double _width = 0;
        private double _height = 0;
        private Margin _padding;
        private Borders _borders;
        private Color _background;
        private VerticalAlignment _vAlign = VerticalAlignment.Top;

        private List<Paragraph> _paragraphs = new List<Paragraph>();
        #endregion

        #region Constructor
        internal TableCell(BD.TableCell cell, int rowIndex, int colIndex, double tableWidth, double colWidth, double charSpace, double lineSpace, B.DocGridType gridType)
        {
            _rowIndex = rowIndex;
            _colIndex = colIndex;
            _colSpan = cell.ColumnSpan;

            /*var cw = cell.GetCellWidth();
            if (cw.Type == CellWidthType.Point)
                _width = cw.Val.ToPixel();
            else if (cw.Type == CellWidthType.Percent)
                _width = tableWidth * cw.Val / 100;
            else
                _width = colWidth;*/

            _width = colWidth;
            foreach (var p in cell.Paragraphs)
            {
                var paragraph  = new Paragraph(p, _width - 12, charSpace, lineSpace, gridType);
                foreach(var line in paragraph.GenerateLines())
                {
                    paragraph.Lines.Add(line);
                }
                _paragraphs.Add(paragraph);
                _height += paragraph.Height;
            }
            _padding = new Margin(6, 0, 6, 0);

            _borders = new Borders();
            _borders.Left.Visible = (int)cell.Borders.Left.Style > 1;
            _borders.Left.Width = cell.Borders.Left.Width.ToPixel();
            _borders.Left.Color = cell.Borders.Left.Color.IsAuto ? Color.Black : cell.Borders.Left.Color.Val;
            _borders.Top.Visible = (int)cell.Borders.Top.Style > 1;
            _borders.Top.Width = cell.Borders.Top.Width.ToPixel();
            _borders.Top.Color = cell.Borders.Top.Color.IsAuto ? Color.Black : cell.Borders.Top.Color.Val;
            _borders.Right.Visible = (int)cell.Borders.Right.Style > 1;
            _borders.Right.Width = cell.Borders.Right.Width.ToPixel();
            _borders.Right.Color = cell.Borders.Right.Color.IsAuto ? Color.Black : cell.Borders.Right.Color.Val;
            _borders.Bottom.Visible = (int)cell.Borders.Bottom.Style > 1;
            _borders.Bottom.Width = cell.Borders.Bottom.Width.ToPixel();
            _borders.Bottom.Color = cell.Borders.Bottom.Color.IsAuto ? Color.Black : cell.Borders.Bottom.Color.Val;

            _background = cell.Background.IsAuto ? Color.White : cell.Background.Val;

            _vAlign = cell.VerticalCellAlignment.Convert<VerticalAlignment>();
        }
        #endregion

        #region Public Properties
        public int RowIndex => _rowIndex;

        public int ColumnIndex => _colIndex;

        public int RowSpan
        {
            get => _rowSpan;
            internal set => _rowSpan = value;
        }

        public int ColumnSpan => _colSpan;

        public double Width => _width;

        public double Height
        {
            get => _height;
            internal set => _height = value;
        }

        public Margin Padding => _padding;

        public Borders Borders => _borders;

        public Color Background => _background;

        public VerticalAlignment VerticalAlignment => _vAlign;

        public List<Paragraph> Paragraphs => _paragraphs;
        #endregion
    }
}
