using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BD = Berry.Docx.Documents;

namespace Berry.Docx.Visual.Documents
{
    public class Table : DocumentItem
    {
        #region Private Members
        private double _width = 0;
        private double _height = 0;
        private List<double> _rowHeights = new List<double>();
        private List<double> _colWidths = new List<double>();
        private List<TableCell> _cells;
        private HorizontalAlignment _hAlign = HorizontalAlignment.Left;
        #endregion

        #region Constructors
        internal Table(BD.Table table, double charSpace, double lineSpace, Berry.Docx.DocGridType gridType)
        {
            _cells = new List<TableCell>();
            int rowIndex = 0;
            foreach(var col in table.ColumnWidths)
            {
                _colWidths.Add(col.Width.ToPixel());
            }
            _width = _colWidths.Sum();
            foreach (var row in table.Rows)
            {
                int colIndex = 0;
                double maxCellHeight = 0;
                foreach (var tc in row.Cells)
                {
                    double colWidth = 0;
                    for(int col = colIndex; col < colIndex + tc.ColumnSpan; col++)
                        colWidth += _colWidths[col];
                    TableCell cell = new TableCell(tc, rowIndex, colIndex, _width, colWidth,
                        charSpace, lineSpace, gridType);
                    if(tc.VerticalMerge == TableCellVerticalMergeType.Continue)
                    {
                        var topCell = _cells.Where(c => c.ColumnIndex == colIndex).LastOrDefault();
                        if(topCell != null) topCell.RowSpan++;
                        colIndex += tc.ColumnSpan;
                        continue;
                    }
                    _cells.Add(cell);
                    colIndex += tc.ColumnSpan;
                    maxCellHeight = Math.Max(maxCellHeight, cell.Height);
                }
                if (row.HeightType == TableRowHeightType.Exactly)
                    _rowHeights.Add(row.Height.ToPixel());
                else
                    _rowHeights.Add(Math.Max(row.Height.ToPixel(), maxCellHeight));
                rowIndex++;
            }

            
            foreach(var cell in _cells)
            {
                cell.Height = 0;
                for (int row = cell.RowIndex; row < cell.RowIndex + cell.RowSpan; row++)
                {
                    cell.Height += _rowHeights[row];
                }
                // Deal border conflict with right cell.
                var rCells = GetRightCells(_cells, cell);
                if (rCells.Count > 0)
                {
                    var rc = rCells[0];
                    if (cell.Borders.Right.Visible && rc.Borders.Left.Visible)
                    {
                        double width = Math.Max(cell.Borders.Right.Width, rc.Borders.Left.Width) / 2;
                        if(cell.Borders.Right.Width > rc.Borders.Left.Width)
                        {
                            foreach (var c in rCells) c.Borders.Left.Color = cell.Borders.Right.Color;
                        }
                        else if(cell.Borders.Right.Width < rc.Borders.Left.Width)
                        {
                            cell.Borders.Right.Color = rc.Borders.Left.Color;
                        }
                        else
                        {
                            int brightness = GetBrightness1(cell.Borders.Right.Color);
                            int brightnessRight = GetBrightness1(rc.Borders.Left.Color);
                            if(brightness < brightnessRight)
                            {
                                foreach (var c in rCells) c.Borders.Left.Color = cell.Borders.Right.Color;
                            }
                            else if(brightness > brightnessRight)
                            {
                                cell.Borders.Right.Color = rc.Borders.Left.Color;
                            }
                            else
                            {
                                brightness = GetBrightness2(cell.Borders.Right.Color);
                                brightnessRight = GetBrightness2(rc.Borders.Left.Color);
                                if (brightness < brightnessRight)
                                {
                                    foreach (var c in rCells) c.Borders.Left.Color = cell.Borders.Right.Color;
                                }
                                else if (brightness > brightnessRight)
                                {
                                    cell.Borders.Right.Color = rc.Borders.Left.Color;
                                }
                                else
                                {
                                    brightness = GetBrightness3(cell.Borders.Right.Color);
                                    brightnessRight = GetBrightness3(rc.Borders.Left.Color);
                                    if (brightness > brightnessRight)
                                    {
                                        cell.Borders.Right.Color = rc.Borders.Left.Color;
                                    }
                                    else
                                    {
                                        foreach (var c in rCells) c.Borders.Left.Color = cell.Borders.Right.Color;
                                    }
                                }
                            }
                        }
                        cell.Borders.Right.Width = width;
                        foreach (var c in rCells) c.Borders.Left.Width = width;
                    }
                    else if(cell.Borders.Right.Visible && !rc.Borders.Left.Visible)
                    {
                        double width = cell.Borders.Right.Width / 2;
                        foreach (var c in rCells)
                        {
                            c.Borders.Left.Width = width;
                            c.Borders.Left.Color = cell.Borders.Right.Color;
                        }
                        cell.Borders.Right.Width = width;
                    }
                    else if (!cell.Borders.Right.Visible && rc.Borders.Left.Visible)
                    {
                        double width = rc.Borders.Right.Width / 2;
                        foreach (var c in rCells)
                        {
                            c.Borders.Left.Width = width;
                        }
                        cell.Borders.Right.Width = width;
                        cell.Borders.Right.Color = rc.Borders.Left.Color;
                    }
                    else
                    {
                        cell.Borders.Right.Width = 0;
                        foreach (var c in rCells) c.Borders.Left.Width = 0;
                    }
                }

                // Deal border conflict with bottom cell.
                var bCells = GetBottomCells(_cells, cell);
                if(bCells.Count > 0)
                {
                    var bc = bCells[0];
                    if (cell.Borders.Bottom.Visible && bc.Borders.Top.Visible)
                    {
                        if(cell.Borders.Bottom.Width > bc.Borders.Top.Width)
                        {
                            foreach (var c in bCells)
                            {
                                c.Borders.Top.Width = cell.Borders.Bottom.Width;
                                c.Borders.Top.Color = cell.Borders.Bottom.Color;
                            }
                            cell.Borders.Bottom.Width = 0;
                        }
                        else if(cell.Borders.Bottom.Width < bc.Borders.Top.Width)
                        {
                            cell.Borders.Bottom.Width = 0;
                        }
                        else
                        {
                            int brightness = GetBrightness1(cell.Borders.Bottom.Color);
                            int brightnessBottom = GetBrightness1(bc.Borders.Top.Color);
                            if(brightness < brightnessBottom)
                            {
                                foreach (var c in bCells) c.Borders.Top.Color = cell.Borders.Bottom.Color;
                            }
                            else if(brightness == brightnessBottom)
                            {
                                brightness = GetBrightness2(cell.Borders.Bottom.Color);
                                brightnessBottom = GetBrightness2(bc.Borders.Top.Color);
                                if (brightness < brightnessBottom)
                                {
                                    foreach (var c in bCells) c.Borders.Top.Color = cell.Borders.Bottom.Color;
                                }
                                else if (brightness == brightnessBottom)
                                {
                                    brightness = GetBrightness3(cell.Borders.Bottom.Color);
                                    brightnessBottom = GetBrightness3(bc.Borders.Top.Color);
                                    if (brightness <= brightnessBottom)
                                    {
                                        foreach (var c in bCells) c.Borders.Top.Color = cell.Borders.Bottom.Color;
                                    }
                                }
                            }
                        }
                    }
                    else if(cell.Borders.Bottom.Visible && !bc.Borders.Top.Visible)
                    {
                        foreach (var c in bCells)
                        {
                            c.Borders.Top.Width = cell.Borders.Bottom.Width;
                            c.Borders.Top.Color = cell.Borders.Bottom.Color;
                        }
                    }
                    else if (!cell.Borders.Bottom.Visible && !bc.Borders.Top.Visible)
                    {
                        foreach (var c in bCells) c.Borders.Top.Width = 0;
                    }
                    cell.Borders.Bottom.Width = 0;
                }
            }

            _height = _rowHeights.Sum();
            _hAlign = table.Format.HorizontalAlignment.Convert<HorizontalAlignment>();
        }
        #endregion

        #region Public Properties
        public double Width => _width;

        public double Height => _height;

        public List<double> RowHeights => _rowHeights;

        public List<double> ColumnWidths => _colWidths;

        public HorizontalAlignment HorizontalAlignment => _hAlign;

        public List<TableCell> Cells => _cells;
        #endregion

        #region Private Methods
        private List<TableCell> GetRightCells(List<TableCell> cells, TableCell cell)
        {
            int colCnt = _colWidths.Count;
            if (cell.ColumnIndex + cell.ColumnSpan == colCnt) return new List<TableCell>();
            return cells.Where(c => 
            c.ColumnIndex == cell.ColumnIndex + cell.ColumnSpan
            && c.RowIndex >= cell.RowIndex 
            && c.RowIndex < cell.RowIndex + cell.RowSpan).ToList();
        }

        private List<TableCell> GetBottomCells(List<TableCell> cells, TableCell cell)
        {
            int rowCnt = _rowHeights.Count;
            if (cell.RowIndex + cell.RowSpan == rowCnt) return new List<TableCell>();
            return cells.Where(c => 
            c.RowIndex == cell.RowIndex + cell.RowSpan
            && c.ColumnIndex >= cell.ColumnIndex 
            && c.ColumnIndex < cell.ColumnIndex + cell.ColumnSpan).ToList();
        }

        private int GetBrightness1(System.Drawing.Color color)
        {
            return color.R + color.B + color.G * 2;
        }
        private int GetBrightness2(System.Drawing.Color color)
        {
            return color.B + color.G * 2;
        }
        private int GetBrightness3(System.Drawing.Color color)
        {
            return color.G * 2;
        }
        #endregion
    }
}
