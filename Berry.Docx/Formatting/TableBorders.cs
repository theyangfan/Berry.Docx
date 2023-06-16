using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represents the table cell borders.
    /// </summary>
    public class TableBorders
    {
        #region Private Members
        private readonly Table _table;
        private readonly TableCell _cell;
        private readonly Style _style;
        private readonly TableRegionType _region;

        private readonly TableBorder _left;
        private readonly TableBorder _top;
        private readonly TableBorder _right;
        private readonly TableBorder _bottom;
        private readonly TableBorder _insideH;
        private readonly TableBorder _insideV;
        #endregion

        #region Constructors
        internal TableBorders(Style style, TableRegionType region)
        {
            _style = style;
            _region = region;
            _left = new TableBorder(style, region, TableBorderType.Left);
            _right = new TableBorder(style, region, TableBorderType.Right);
            _top = new TableBorder(style, region, TableBorderType.Top);
            _bottom = new TableBorder(style, region, TableBorderType.Bottom);
            _insideH = new TableBorder(style, region, TableBorderType.InsideH);
            _insideV = new TableBorder(style, region, TableBorderType.InsideV);
        }

        internal TableBorders(Table table)
        {
            _table = table;
            _left = new TableBorder(table, TableBorderType.Left);
            _right = new TableBorder(table, TableBorderType.Right);
            _top = new TableBorder(table, TableBorderType.Top);
            _bottom = new TableBorder(table, TableBorderType.Bottom);
            _insideH = new TableBorder(table, TableBorderType.InsideH);
            _insideV = new TableBorder(table, TableBorderType.InsideV);
        }

        internal TableBorders(TableCell cell)
        {
            _cell = cell;
            _left = new TableBorder(cell, TableBorderType.Left);
            _right = new TableBorder(cell, TableBorderType.Right);
            _top = new TableBorder(cell, TableBorderType.Top);
            _bottom = new TableBorder(cell, TableBorderType.Bottom);
            _insideH = new TableBorder(cell, TableBorderType.InsideH);
            _insideV = new TableBorder(cell, TableBorderType.InsideV);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the table cell top border.
        /// </summary>
        public TableBorder Top => _top;

        /// <summary>
        /// Gets the table cell bottom border.
        /// </summary>
        public TableBorder Bottom => _bottom;

        /// <summary>
        /// Gets the table cell left border.
        /// </summary>
        public TableBorder Left => _left;

        /// <summary>
        /// Gets the table cell right border.
        /// </summary>
        public TableBorder Right => _right;

        /// <summary>
        /// Gets the table cell inside horizontal border.
        /// </summary>
        public TableBorder InsideH => _insideH;

        /// <summary>
        /// Gets the table cell inside vertical border.
        /// </summary>
        public TableBorder InsideV => _insideV;
        #endregion

        #region Public Methods
        /// <summary>
        /// Clears all table borders.
        /// </summary>
        public void Clear()
        {
            if(_cell != null)
            {
                _cell.XElement?.TableCellProperties?.TableCellBorders?.Remove();
            }
            else if(_table != null)
            {
                _table.XElement.GetFirstChild<W.TableProperties>()?.TableBorders?.Remove();
            }
            else if(_style != null)
            {
                if (_region == TableRegionType.WholeTable)
                {
                    _style.XElement.StyleTableProperties?.TableBorders?.Remove();
                }
                else
                {
                    var type = _region.Convert<W.TableStyleOverrideValues>();
                    var tblStylePr = _style.XElement.Elements<W.TableStyleProperties>()
                        .Where(t => t.Type == type).FirstOrDefault();
                    tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.Remove();
                }
            }
        }
        #endregion

    }

    /// <summary>
    /// Represents the table cell border.
    /// </summary>
    public class TableBorder
    {
        #region Private Members
        private readonly Table _table;
        private readonly TableCell _cell;
        private readonly Style _style;
        private readonly TableRegionType _region;
        private readonly TableBorderType _type;
        #endregion

        #region Constructors
        internal TableBorder(TableCell cell, TableBorderType type)
        {
            _cell = cell;
            _type = type;
        }

        internal TableBorder(Table table, TableBorderType type)
        {
            _table = table;
            _type = type;
        }

        internal TableBorder(Style style, TableRegionType region, TableBorderType type)
        {
            _style = style;
            _region = region;
            _type = type;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets or sets the border style.
        /// </summary>
        public BorderStyle Style
        {
            get
            {
                if (_cell != null)
                {
                    var cellBorder = new TableBorderHolder(_cell.XElement, _type);
                    if (cellBorder.Style != null)
                    {
                        return cellBorder.Style;
                    }
                    else if(_cell.Table != null)
                    {
                        var tblBdrs = new TableBorders(_cell.Table);
                        int rowCnt = _cell.Table.RowCount;
                        int colCnt = _cell.Table.ColumnCount;
                        int rowIndex = _cell.RowIndex;
                        int colIndex = _cell.ColumnIndex;
                        int colSpan = _cell.ColumnSpan;
                        if (_type == TableBorderType.Left)
                        {
                            if (colIndex == 0) return tblBdrs.Left.Style;
                            else return tblBdrs.InsideV.Style;
                        }
                        else if(_type == TableBorderType.Top)
                        {
                            if(rowIndex == 0) return tblBdrs.Top.Style;
                            else return tblBdrs.InsideH.Style;
                        }
                        else if (_type == TableBorderType.Right)
                        {
                            if (colIndex + colSpan == colCnt) return tblBdrs.Right.Style;
                            else return tblBdrs.InsideV.Style;
                        }
                        else if (_type == TableBorderType.Bottom)
                        {
                            if (rowIndex == rowCnt - 1) return tblBdrs.Bottom.Style;
                            else return tblBdrs.InsideH.Style;
                        }
                    }
                }
                else if (_table != null)
                {
                    var tblBorder = new TableBorderHolder(_table.XElement, _type);
                    if (tblBorder.Style != null) return tblBorder.Style;
                    var style = _table.GetStyle();
                    if (style != null)
                    {
                        return new TableBorder(style, TableRegionType.WholeTable, _type).Style;
                    }
                }
                else if(_style != null)
                {
                    var styleBdr = new TableBorderHolder(_style.XElement, _region, _type);
                    var styleWholeTblBdr = new TableBorderHolder(_style.XElement, TableRegionType.WholeTable, _type);
                    if (styleBdr.Style != null) 
                        return styleBdr.Style;
                    if(_region != TableRegionType.WholeTable && styleWholeTblBdr.Style != null) 
                        return styleWholeTblBdr.Style;
                    if (_style.BaseStyle != null)
                        return new TableBorder(_style.BaseStyle, _region, _type).Style;
                }
                return BorderStyle.Nil;
            }
            set
            {
                if(_cell != null)
                {
                    var border = new TableBorderHolder(_cell.XElement, _type);
                    border.Style = value;
                }
                else if(_table != null)
                {
                    var border = new TableBorderHolder(_table.XElement, _type);
                    border.Style = value;
                }
                else if(_style != null)
                {
                    var border = new TableBorderHolder(_style.XElement, _region, _type);
                    border.Style = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the border width in points.
        /// </summary>
        public float Width
        {
            get
            {
                if (_cell != null)
                {
                    var cellBorder = new TableBorderHolder(_cell.XElement, _type);
                    if (cellBorder.Width != null)
                    {
                        return cellBorder.Width;
                    }
                    else if (_cell.Table != null)
                    {
                        var tblBdrs = new TableBorders(_cell.Table);
                        int rowCnt = _cell.Table.RowCount;
                        int colCnt = _cell.Table.ColumnCount;
                        int rowIndex = _cell.RowIndex;
                        int colIndex = _cell.ColumnIndex;
                        int colSpan = _cell.ColumnSpan;
                        if (_type == TableBorderType.Left)
                        {
                            if (colIndex == 0) return tblBdrs.Left.Width;
                            else return tblBdrs.InsideV.Width;
                        }
                        else if (_type == TableBorderType.Top)
                        {
                            if (rowIndex == 0) return tblBdrs.Top.Width;
                            else return tblBdrs.InsideH.Width;
                        }
                        else if (_type == TableBorderType.Right)
                        {
                            if (colIndex + colSpan == colCnt) return tblBdrs.Right.Width;
                            else return tblBdrs.InsideV.Width;
                        }
                        else if (_type == TableBorderType.Bottom)
                        {
                            if (rowIndex == rowCnt - 1) return tblBdrs.Bottom.Width;
                            else return tblBdrs.InsideH.Width;
                        }
                    }
                }
                else if (_table != null)
                {
                    var tblBorder = new TableBorderHolder(_table.XElement, _type);
                    if (tblBorder.Width != null) return tblBorder.Width;
                    var style = _table.GetStyle();
                    if (style != null)
                    {
                        return new TableBorder(style, TableRegionType.WholeTable, _type).Width;
                    }
                }
                else if (_style != null)
                {
                    var styleBdr = new TableBorderHolder(_style.XElement, _region, _type);
                    var styleWholeTblBdr = new TableBorderHolder(_style.XElement, TableRegionType.WholeTable, _type);
                    if (styleBdr.Width != null)
                        return styleBdr.Width;
                    if (_region != TableRegionType.WholeTable && styleWholeTblBdr.Width != null)
                        return styleWholeTblBdr.Width;
                    if (_style.BaseStyle != null)
                        return new TableBorder(_style.BaseStyle, _region, _type).Width;
                }
                return 0;
            }
            set
            {
                if (_cell != null)
                {
                    var border = new TableBorderHolder(_cell.XElement, _type);
                    border.Width = value;
                }
                else if (_table != null)
                {
                    var border = new TableBorderHolder(_table.XElement, _type);
                    border.Width = value;
                }
                else if (_style != null)
                {
                    var border = new TableBorderHolder(_style.XElement, _region, _type);
                    border.Width = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the border color.
        /// </summary>
        public ColorValue Color
        {
            get
            {
                if (_cell != null)
                {
                    var cellBorder = new TableBorderHolder(_cell.XElement, _type);
                    if (cellBorder.Color != null)
                    {
                        return cellBorder.Color;
                    }
                    else if (_cell.Table != null)
                    {
                        var tblBdrs = new TableBorders(_cell.Table);
                        int rowCnt = _cell.Table.RowCount;
                        int colCnt = _cell.Table.ColumnCount;
                        int rowIndex = _cell.RowIndex;
                        int colIndex = _cell.ColumnIndex;
                        int colSpan = _cell.ColumnSpan;
                        if (_type == TableBorderType.Left)
                        {
                            if (colIndex == 0) return tblBdrs.Left.Color;
                            else return tblBdrs.InsideV.Color;
                        }
                        else if (_type == TableBorderType.Top)
                        {
                            if (rowIndex == 0) return tblBdrs.Top.Color;
                            else return tblBdrs.InsideH.Color;
                        }
                        else if (_type == TableBorderType.Right)
                        {
                            if (colIndex + colSpan == colCnt) return tblBdrs.Right.Color;
                            else return tblBdrs.InsideV.Color;
                        }
                        else if (_type == TableBorderType.Bottom)
                        {
                            if (rowIndex == rowCnt - 1) return tblBdrs.Bottom.Color;
                            else return tblBdrs.InsideH.Color;
                        }
                    }
                }
                else if (_table != null)
                {
                    var tblBorder = new TableBorderHolder(_table.XElement, _type);
                    if (tblBorder.Color != null) return tblBorder.Color;
                    var style = _table.GetStyle();
                    if (style != null)
                    {
                        return new TableBorder(style, TableRegionType.WholeTable, _type).Color;
                    }
                }
                else if (_style != null)
                {
                    var styleBdr = new TableBorderHolder(_style.XElement, _region, _type);
                    var styleWholeTblBdr = new TableBorderHolder(_style.XElement, TableRegionType.WholeTable, _type);
                    if (styleBdr.Color != null)
                        return styleBdr.Color;
                    if (_region != TableRegionType.WholeTable && styleWholeTblBdr.Color != null)
                        return styleWholeTblBdr.Color;
                    if (_style.BaseStyle != null)
                        return new TableBorder(_style.BaseStyle, _region, _type).Color;
                }
                return ColorValue.Auto;
            }
            set
            {
                if (_cell != null)
                {
                    var border = new TableBorderHolder(_cell.XElement, _type);
                    border.Color = value;
                }
                else if (_table != null)
                {
                    var border = new TableBorderHolder(_table.XElement, _type);
                    border.Color = value;
                }
                else if (_style != null)
                {
                    var border = new TableBorderHolder(_style.XElement, _region, _type);
                    border.Color = value;
                }
            }
        }
        #endregion
    }

    internal class TableBorderHolder
    {
        private readonly W.Table _table;
        private readonly W.TableCell _cell;
        private readonly W.Style _style;
        private readonly TableRegionType _styleRegion;
        private readonly TableBorderType _type;

        internal TableBorderHolder(W.TableCell cell, TableBorderType type)
        {
            _cell = cell;
            _type = type;
        }

        internal TableBorderHolder(W.Table table, TableBorderType type)
        {
            _table = table;
            _type = type;
        }
        internal TableBorderHolder(W.Style style, TableRegionType region, TableBorderType type)
        {
            _style = style;
            _styleRegion = region;
            _type = type;
        }

        public EnumValue<BorderStyle> Style
        {
            get
            {
                var border = TryGetBorder(_type);
                if (border?.Val == null) return null;
                return border.Val.Value.Convert<BorderStyle>();
            }
            set
            {
                CreateBorder(out W.BorderType border);
                border.Val = value.Val.Convert<W.BorderValues>();
            }
        }

        public FloatValue Width
        {
            get
            {
                if(Style == null) return null;
                var border = TryGetBorder(_type);
                if (border?.Size == null) return null;
                if ((int)Style.Val < 27)
                    return border.Size.Value / 8.0F;
                else
                    return border.Size.Value;
            }
            set
            {
                CreateBorder(out W.BorderType border);
                if (Style == null || (int)Style.Val < 27)
                {
                    if (value > 12)
                        border.Size = 96;
                    else if (value >= 0.25)
                        border.Size = (uint)(value * 8);
                    else if (value > 0)
                        border.Size = 2;
                    else
                        border.Size = 0;
                }
                else
                {
                    if (value > 31)
                        border.Size = 31;
                    else if (value >= 1)
                        border.Size = (uint)value;
                    else if (value > 0)
                        border.Size = 1;
                    else
                        border.Size = 0;
                }
            }
        }

        public ColorValue Color
        {
            get
            {
                if (Style == null) return null;
                var border = TryGetBorder(_type);
                if (border?.Color == null) return null;
                return border.Color.Value;
            }
            set
            {
                CreateBorder(out W.BorderType border);
                border.Color = value.ToString();
            }
        }

        private W.BorderType TryGetBorder(TableBorderType borderType)
        {
            W.BorderType border = null;
            switch (borderType)
            {
                case TableBorderType.Left:
                    if (_cell != null)
                    {
                        border = _cell.TableCellProperties?.TableCellBorders?.LeftBorder;
                    }
                    else if (_table != null)
                    {
                        border = _table.GetFirstChild<W.TableProperties>()?.TableBorders?.LeftBorder;
                    }
                    else if (_style != null)
                    {
                        if (_styleRegion == TableRegionType.WholeTable)
                        {
                            border = _style.StyleTableProperties?.TableBorders?.LeftBorder;
                        }
                        else
                        {
                            W.TableStyleOverrideValues type = _styleRegion.Convert<W.TableStyleOverrideValues>();
                            W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == type).FirstOrDefault();
                            border = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.LeftBorder;
                        }
                    }
                    break;
                case TableBorderType.Top:
                    if (_cell != null)
                    {
                        border = _cell.TableCellProperties?.TableCellBorders?.TopBorder;
                    }
                    else if (_table != null)
                    {
                        border = _table.GetFirstChild<W.TableProperties>()?.TableBorders?.TopBorder;
                    }
                    else if (_style != null)
                    {
                        if (_styleRegion == TableRegionType.WholeTable)
                        {
                            border = _style.StyleTableProperties?.TableBorders?.TopBorder;
                        }
                        else
                        {
                            W.TableStyleOverrideValues type = _styleRegion.Convert<W.TableStyleOverrideValues>();
                            W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == type).FirstOrDefault();
                            border = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.TopBorder;
                        }
                    }
                    break;
                case TableBorderType.Right:
                    if (_cell != null)
                    {
                        border = _cell.TableCellProperties?.TableCellBorders?.RightBorder;
                    }
                    else if (_table != null)
                    {
                        border = _table.GetFirstChild<W.TableProperties>()?.TableBorders?.RightBorder;
                    }
                    else if (_style != null)
                    {
                        if (_styleRegion == TableRegionType.WholeTable)
                        {
                            border = _style.StyleTableProperties?.TableBorders?.RightBorder;
                        }
                        else
                        {
                            W.TableStyleOverrideValues type = _styleRegion.Convert<W.TableStyleOverrideValues>();
                            W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == type).FirstOrDefault();
                            border = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.RightBorder;
                        }
                    }
                    break;
                case TableBorderType.Bottom:
                    if (_cell != null)
                    {
                        border = _cell.TableCellProperties?.TableCellBorders?.BottomBorder;
                    }
                    else if (_table != null)
                    {
                        border = _table.GetFirstChild<W.TableProperties>()?.TableBorders?.BottomBorder;
                    }
                    else if (_style != null)
                    {
                        if (_styleRegion == TableRegionType.WholeTable)
                        {
                            border = _style.StyleTableProperties?.TableBorders?.BottomBorder;
                        }
                        else
                        {
                            W.TableStyleOverrideValues type = _styleRegion.Convert<W.TableStyleOverrideValues>();
                            W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == type).FirstOrDefault();
                            border = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.BottomBorder;
                        }
                    }
                    break;
                case TableBorderType.InsideH:
                    if (_cell != null)
                    {
                        border = _cell.TableCellProperties?.TableCellBorders?.InsideHorizontalBorder;
                    }
                    else if (_table != null)
                    {
                        border = _table.GetFirstChild<W.TableProperties>()?.TableBorders?.InsideHorizontalBorder;
                    }
                    else if (_style != null)
                    {
                        if (_styleRegion == TableRegionType.WholeTable)
                        {
                            border = _style.StyleTableProperties?.TableBorders?.InsideHorizontalBorder;
                        }
                        else
                        {
                            W.TableStyleOverrideValues type = _styleRegion.Convert<W.TableStyleOverrideValues>();
                            W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == type).FirstOrDefault();
                            border = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.InsideHorizontalBorder;
                        }
                    }
                    break;
                case TableBorderType.InsideV:
                    if (_cell != null)
                    {
                        border = _cell.TableCellProperties?.TableCellBorders?.InsideVerticalBorder;
                    }
                    else if (_table != null)
                    {
                        border = _table.GetFirstChild<W.TableProperties>()?.TableBorders?.InsideVerticalBorder;
                    }
                    else if (_style != null)
                    {
                        if (_styleRegion == TableRegionType.WholeTable)
                        {
                            border = _style.StyleTableProperties?.TableBorders?.InsideVerticalBorder;
                        }
                        else
                        {
                            W.TableStyleOverrideValues type = _styleRegion.Convert<W.TableStyleOverrideValues>();
                            W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == type).FirstOrDefault();
                            border = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.InsideVerticalBorder;
                        }
                    }
                    break;
                default:
                    break;
            }
            return border;
        }
        private void CreateBorder(out W.BorderType border)
        {
            border = null;
            switch (_type)
            {
                case TableBorderType.Left:
                    if(_cell != null)
                    {
                        if(_cell.TableCellProperties == null)
                            _cell.TableCellProperties = new W.TableCellProperties();
                        if(_cell.TableCellProperties.TableCellBorders == null)
                            _cell.TableCellProperties.TableCellBorders = new W.TableCellBorders();
                        if(_cell.TableCellProperties.TableCellBorders.LeftBorder == null)
                            _cell.TableCellProperties.TableCellBorders.LeftBorder = new W.LeftBorder();
                        border = _cell.TableCellProperties.TableCellBorders.LeftBorder;
                    }
                    else if(_table != null)
                    {
                        if (_table.GetFirstChild<W.TableProperties>() == null)
                            _table.AddChild(new W.TableProperties());
                        if(_table.GetFirstChild<W.TableProperties>().TableBorders == null)
                            _table.GetFirstChild<W.TableProperties>().TableBorders = new W.TableBorders();
                        if (_table.GetFirstChild<W.TableProperties>().TableBorders.LeftBorder == null)
                            _table.GetFirstChild<W.TableProperties>().TableBorders.LeftBorder = new W.LeftBorder();
                        border = _table.GetFirstChild<W.TableProperties>().TableBorders.LeftBorder;
                    }
                    else if(_style != null)
                    {
                        if(_styleRegion == TableRegionType.WholeTable)
                        {
                            if (_style.StyleTableProperties == null)
                                _style.StyleTableProperties = new W.StyleTableProperties();
                            if (_style.StyleTableProperties.TableBorders == null)
                                _style.StyleTableProperties.TableBorders = new W.TableBorders();
                            if (_style.StyleTableProperties.TableBorders.LeftBorder == null)
                                _style.StyleTableProperties.TableBorders.LeftBorder = new W.LeftBorder();
                            border = _style.StyleTableProperties.TableBorders.LeftBorder;
                        }
                        else
                        {
                            W.TableStyleOverrideValues type = _styleRegion.Convert<W.TableStyleOverrideValues>();
                            if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                            {
                                _style.Append(new W.TableStyleProperties() { Type = type });
                            }
                            W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties = new W.TableStyleConditionalFormattingTableCellProperties();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders = new W.TableCellBorders();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.LeftBorder == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.LeftBorder = new W.LeftBorder();
                            border = tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.LeftBorder;
                        }
                    }
                    break;
                case TableBorderType.Top:
                    if (_cell != null)
                    {
                        if (_cell.TableCellProperties == null)
                            _cell.TableCellProperties = new W.TableCellProperties();
                        if (_cell.TableCellProperties.TableCellBorders == null)
                            _cell.TableCellProperties.TableCellBorders = new W.TableCellBorders();
                        if (_cell.TableCellProperties.TableCellBorders.TopBorder == null)
                            _cell.TableCellProperties.TableCellBorders.TopBorder = new W.TopBorder();
                        border = _cell.TableCellProperties.TableCellBorders.TopBorder;
                    }
                    else if (_table != null)
                    {
                        if (_table.GetFirstChild<W.TableProperties>() == null)
                            _table.AddChild(new W.TableProperties());
                        if (_table.GetFirstChild<W.TableProperties>().TableBorders == null)
                            _table.GetFirstChild<W.TableProperties>().TableBorders = new W.TableBorders();
                        if (_table.GetFirstChild<W.TableProperties>().TableBorders.TopBorder == null)
                            _table.GetFirstChild<W.TableProperties>().TableBorders.TopBorder = new W.TopBorder();
                        border = _table.GetFirstChild<W.TableProperties>().TableBorders.TopBorder;
                    }
                    else if (_style != null)
                    {
                        if (_styleRegion == TableRegionType.WholeTable)
                        {
                            if (_style.StyleTableProperties == null)
                                _style.StyleTableProperties = new W.StyleTableProperties();
                            if (_style.StyleTableProperties.TableBorders == null)
                                _style.StyleTableProperties.TableBorders = new W.TableBorders();
                            if (_style.StyleTableProperties.TableBorders.TopBorder == null)
                                _style.StyleTableProperties.TableBorders.TopBorder = new W.TopBorder();
                            border = _style.StyleTableProperties.TableBorders.TopBorder;
                        }
                        else
                        {
                            W.TableStyleOverrideValues type = _styleRegion.Convert<W.TableStyleOverrideValues>();
                            if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                            {
                                _style.Append(new W.TableStyleProperties() { Type = type });
                            }
                            W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties = new W.TableStyleConditionalFormattingTableCellProperties();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders = new W.TableCellBorders();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.TopBorder == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.TopBorder = new W.TopBorder();
                            border = tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.TopBorder;
                        }
                    }
                    break;
                case TableBorderType.Right:
                    if (_cell != null)
                    {
                        if (_cell.TableCellProperties == null)
                            _cell.TableCellProperties = new W.TableCellProperties();
                        if (_cell.TableCellProperties.TableCellBorders == null)
                            _cell.TableCellProperties.TableCellBorders = new W.TableCellBorders();
                        if (_cell.TableCellProperties.TableCellBorders.RightBorder == null)
                            _cell.TableCellProperties.TableCellBorders.RightBorder = new W.RightBorder();
                        border = _cell.TableCellProperties.TableCellBorders.RightBorder;
                    }
                    else if (_table != null)
                    {
                        if (_table.GetFirstChild<W.TableProperties>() == null)
                            _table.AddChild(new W.TableProperties());
                        if (_table.GetFirstChild<W.TableProperties>().TableBorders == null)
                            _table.GetFirstChild<W.TableProperties>().TableBorders = new W.TableBorders();
                        if (_table.GetFirstChild<W.TableProperties>().TableBorders.RightBorder == null)
                            _table.GetFirstChild<W.TableProperties>().TableBorders.RightBorder = new W.RightBorder();
                        border = _table.GetFirstChild<W.TableProperties>().TableBorders.RightBorder;
                    }
                    else if (_style != null)
                    {
                        if (_styleRegion == TableRegionType.WholeTable)
                        {
                            if (_style.StyleTableProperties == null)
                                _style.StyleTableProperties = new W.StyleTableProperties();
                            if (_style.StyleTableProperties.TableBorders == null)
                                _style.StyleTableProperties.TableBorders = new W.TableBorders();
                            if (_style.StyleTableProperties.TableBorders.RightBorder == null)
                                _style.StyleTableProperties.TableBorders.RightBorder = new W.RightBorder();
                            border = _style.StyleTableProperties.TableBorders.RightBorder;
                        }
                        else
                        {
                            W.TableStyleOverrideValues type = _styleRegion.Convert<W.TableStyleOverrideValues>();
                            if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                            {
                                _style.Append(new W.TableStyleProperties() { Type = type });
                            }
                            W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties = new W.TableStyleConditionalFormattingTableCellProperties();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders = new W.TableCellBorders();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.RightBorder == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.RightBorder = new W.RightBorder();
                            border = tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.RightBorder;
                        }
                    }
                    break;
                case TableBorderType.Bottom:
                    if (_cell != null)
                    {
                        if (_cell.TableCellProperties == null)
                            _cell.TableCellProperties = new W.TableCellProperties();
                        if (_cell.TableCellProperties.TableCellBorders == null)
                            _cell.TableCellProperties.TableCellBorders = new W.TableCellBorders();
                        if (_cell.TableCellProperties.TableCellBorders.BottomBorder == null)
                            _cell.TableCellProperties.TableCellBorders.BottomBorder = new W.BottomBorder();
                        border = _cell.TableCellProperties.TableCellBorders.BottomBorder;
                    }
                    else if (_table != null)
                    {
                        if (_table.GetFirstChild<W.TableProperties>() == null)
                            _table.AddChild(new W.TableProperties());
                        if (_table.GetFirstChild<W.TableProperties>().TableBorders == null)
                            _table.GetFirstChild<W.TableProperties>().TableBorders = new W.TableBorders();
                        if (_table.GetFirstChild<W.TableProperties>().TableBorders.BottomBorder == null)
                            _table.GetFirstChild<W.TableProperties>().TableBorders.BottomBorder = new W.BottomBorder();
                        border = _table.GetFirstChild<W.TableProperties>().TableBorders.BottomBorder;
                    }
                    else if (_style != null)
                    {
                        if (_styleRegion == TableRegionType.WholeTable)
                        {
                            if (_style.StyleTableProperties == null)
                                _style.StyleTableProperties = new W.StyleTableProperties();
                            if (_style.StyleTableProperties.TableBorders == null)
                                _style.StyleTableProperties.TableBorders = new W.TableBorders();
                            if (_style.StyleTableProperties.TableBorders.BottomBorder == null)
                                _style.StyleTableProperties.TableBorders.BottomBorder = new W.BottomBorder();
                            border = _style.StyleTableProperties.TableBorders.BottomBorder;
                        }
                        else
                        {
                            W.TableStyleOverrideValues type = _styleRegion.Convert<W.TableStyleOverrideValues>();
                            if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                            {
                                _style.Append(new W.TableStyleProperties() { Type = type });
                            }
                            W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties = new W.TableStyleConditionalFormattingTableCellProperties();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders = new W.TableCellBorders();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.BottomBorder == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.BottomBorder = new W.BottomBorder();
                            border = tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.BottomBorder;
                        }
                    }
                    break;
                case TableBorderType.InsideH:
                    if (_cell != null)
                    {
                        if (_cell.TableCellProperties == null)
                            _cell.TableCellProperties = new W.TableCellProperties();
                        if (_cell.TableCellProperties.TableCellBorders == null)
                            _cell.TableCellProperties.TableCellBorders = new W.TableCellBorders();
                        if (_cell.TableCellProperties.TableCellBorders.InsideHorizontalBorder == null)
                            _cell.TableCellProperties.TableCellBorders.InsideHorizontalBorder = new W.InsideHorizontalBorder();
                        border = _cell.TableCellProperties.TableCellBorders.InsideHorizontalBorder;
                    }
                    else if (_table != null)
                    {
                        if (_table.GetFirstChild<W.TableProperties>() == null)
                            _table.AddChild(new W.TableProperties());
                        if (_table.GetFirstChild<W.TableProperties>().TableBorders == null)
                            _table.GetFirstChild<W.TableProperties>().TableBorders = new W.TableBorders();
                        if (_table.GetFirstChild<W.TableProperties>().TableBorders.InsideHorizontalBorder == null)
                            _table.GetFirstChild<W.TableProperties>().TableBorders.InsideHorizontalBorder = new W.InsideHorizontalBorder();
                        border = _table.GetFirstChild<W.TableProperties>().TableBorders.InsideHorizontalBorder;
                    }
                    else if (_style != null)
                    {
                        if (_styleRegion == TableRegionType.WholeTable)
                        {
                            if (_style.StyleTableProperties == null)
                                _style.StyleTableProperties = new W.StyleTableProperties();
                            if (_style.StyleTableProperties.TableBorders == null)
                                _style.StyleTableProperties.TableBorders = new W.TableBorders();
                            if (_style.StyleTableProperties.TableBorders.InsideHorizontalBorder == null)
                                _style.StyleTableProperties.TableBorders.InsideHorizontalBorder = new W.InsideHorizontalBorder();
                            border = _style.StyleTableProperties.TableBorders.InsideHorizontalBorder;
                        }
                        else
                        {
                            W.TableStyleOverrideValues type = _styleRegion.Convert<W.TableStyleOverrideValues>();
                            if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                            {
                                _style.Append(new W.TableStyleProperties() { Type = type });
                            }
                            W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties = new W.TableStyleConditionalFormattingTableCellProperties();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders = new W.TableCellBorders();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideHorizontalBorder == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideHorizontalBorder = new W.InsideHorizontalBorder();
                            border = tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideHorizontalBorder;
                        }
                    }
                    break;
                case TableBorderType.InsideV:
                    if (_cell != null)
                    {
                        if (_cell.TableCellProperties == null)
                            _cell.TableCellProperties = new W.TableCellProperties();
                        if (_cell.TableCellProperties.TableCellBorders == null)
                            _cell.TableCellProperties.TableCellBorders = new W.TableCellBorders();
                        if (_cell.TableCellProperties.TableCellBorders.InsideVerticalBorder == null)
                            _cell.TableCellProperties.TableCellBorders.InsideVerticalBorder = new W.InsideVerticalBorder();
                        border = _cell.TableCellProperties.TableCellBorders.InsideVerticalBorder;
                    }
                    else if (_table != null)
                    {
                        if (_table.GetFirstChild<W.TableProperties>() == null)
                            _table.AddChild(new W.TableProperties());
                        if (_table.GetFirstChild<W.TableProperties>().TableBorders == null)
                            _table.GetFirstChild<W.TableProperties>().TableBorders = new W.TableBorders();
                        if (_table.GetFirstChild<W.TableProperties>().TableBorders.InsideVerticalBorder == null)
                            _table.GetFirstChild<W.TableProperties>().TableBorders.InsideVerticalBorder = new W.InsideVerticalBorder();
                        border = _table.GetFirstChild<W.TableProperties>().TableBorders.InsideVerticalBorder;
                    }
                    else if (_style != null)
                    {
                        if (_styleRegion == TableRegionType.WholeTable)
                        {
                            if (_style.StyleTableProperties == null)
                                _style.StyleTableProperties = new W.StyleTableProperties();
                            if (_style.StyleTableProperties.TableBorders == null)
                                _style.StyleTableProperties.TableBorders = new W.TableBorders();
                            if (_style.StyleTableProperties.TableBorders.InsideVerticalBorder == null)
                                _style.StyleTableProperties.TableBorders.InsideVerticalBorder = new W.InsideVerticalBorder();
                            border = _style.StyleTableProperties.TableBorders.InsideVerticalBorder;
                        }
                        else
                        {
                            W.TableStyleOverrideValues type = _styleRegion.Convert<W.TableStyleOverrideValues>();
                            if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                            {
                                _style.Append(new W.TableStyleProperties() { Type = type });
                            }
                            W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties = new W.TableStyleConditionalFormattingTableCellProperties();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders = new W.TableCellBorders();
                            if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideVerticalBorder == null)
                                tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideVerticalBorder = new W.InsideVerticalBorder();
                            border = tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideVerticalBorder;
                        }
                    }
                    break;
                default:
                    break;
            }
        }
    }

    internal enum TableBorderType
    {
        Left = 0,
        Top = 1,
        Right = 2,
        Bottom = 3,
        InsideH = 4,
        InsideV = 5
    }
}
