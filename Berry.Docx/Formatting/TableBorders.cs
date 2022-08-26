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
        private readonly W.Style _style;
        private readonly Table _table;
        private readonly TableCell _cell;

        private readonly TableBorder _left;
        private readonly TableBorder _top;
        private readonly TableBorder _right;
        private readonly TableBorder _bottom;
        private readonly TableBorder _insideH;
        private readonly TableBorder _insideV;
        #endregion

        #region Constructors
        internal TableBorders(Document doc, W.Style style, TableRegionType region)
        {
            _style = style;
            _left = new TableBorder(doc, style, region, TableCellBorderType.Left);
            _right = new TableBorder(doc, style, region, TableCellBorderType.Right);
            _top = new TableBorder(doc, style, region, TableCellBorderType.Top);
            _bottom = new TableBorder(doc, style, region, TableCellBorderType.Bottom);
            _insideH = new TableBorder(doc, style, region, TableCellBorderType.InsideH);
            _insideV = new TableBorder(doc, style, region, TableCellBorderType.InsideV);
        }

        internal TableBorders(Document doc, Table table)
        {
            _table = table;
        }

        internal TableBorders(Document doc, TableCell cell)
        {
            _cell = cell;
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
                if(_cell.XElement?.TableCellProperties?.TableCellBorders != null)
                    _cell.XElement.TableCellProperties.TableCellBorders = null;
            }
            else if(_table != null)
            {
                W.TableProperties tblPr = _table.XElement.GetFirstChild<W.TableProperties>();
                if(tblPr.TableBorders != null)
                {
                    tblPr.TableBorders = null;
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
        private readonly Document _doc;
        private readonly W.Style _style;
        private readonly TableRegionType _region;
        private readonly TableCellBorderType _type;
        #endregion

        #region Constructors
        internal TableBorder(Document doc, W.Style style, TableRegionType region, TableCellBorderType type)
        {
            _doc = doc;
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
                TryGetBorder(out W.BorderType border);
                if (border == null)
                {
                    if (_region == TableRegionType.WholeTable)
                    {
                        W.Style baseStyle = _style.GetBaseStyle(_doc);
                        if (baseStyle != null)
                        {
                            return new TableBorder(_doc, baseStyle, _region, _type).Style;
                        }
                        return BorderStyle.Nil;
                    }
                    else
                    {
                        return new TableBorder(_doc, _style, TableRegionType.WholeTable, _type).Style;
                    }
                }
                if (border.Val == null) return BorderStyle.Nil;
                return border.Val.Value.Convert<BorderStyle>();
            }
            set
            {
                CreateBorder();
                if (_region == TableRegionType.WholeTable)
                {
                    if(_type == TableCellBorderType.Left)
                        _style.StyleTableProperties.TableBorders.LeftBorder.Val = value.Convert<W.BorderValues>();
                    else if (_type == TableCellBorderType.Top)
                        _style.StyleTableProperties.TableBorders.TopBorder.Val = value.Convert<W.BorderValues>();
                    else if (_type == TableCellBorderType.Right)
                        _style.StyleTableProperties.TableBorders.RightBorder.Val = value.Convert<W.BorderValues>();
                    else if (_type == TableCellBorderType.Bottom)
                        _style.StyleTableProperties.TableBorders.BottomBorder.Val = value.Convert<W.BorderValues>();
                    else if (_type == TableCellBorderType.InsideH)
                        _style.StyleTableProperties.TableBorders.InsideHorizontalBorder.Val = value.Convert<W.BorderValues>();
                    else if (_type == TableCellBorderType.InsideV)
                        _style.StyleTableProperties.TableBorders.InsideVerticalBorder.Val = value.Convert<W.BorderValues>();
                }
                else
                {
                    W.TableStyleOverrideValues type = _region.Convert<W.TableStyleOverrideValues>();
                    W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                    if (_type == TableCellBorderType.Left)
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.LeftBorder.Val = value.Convert<W.BorderValues>();
                    else if (_type == TableCellBorderType.Top)
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.TopBorder.Val = value.Convert<W.BorderValues>();
                    else if (_type == TableCellBorderType.Right)
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.RightBorder.Val = value.Convert<W.BorderValues>();
                    else if (_type == TableCellBorderType.Bottom)
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.BottomBorder.Val = value.Convert<W.BorderValues>();
                    else if (_type == TableCellBorderType.InsideH)
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideHorizontalBorder.Val = value.Convert<W.BorderValues>();
                    else if (_type == TableCellBorderType.InsideV)
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideVerticalBorder.Val = value.Convert<W.BorderValues>();
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
                TryGetBorder(out W.BorderType border);
                if (border == null)
                {
                    if (_region == TableRegionType.WholeTable)
                    {
                        W.Style baseStyle = _style.GetBaseStyle(_doc);
                        if (baseStyle != null)
                        {
                            return new TableBorder(_doc, baseStyle, _region, _type).Color;
                        }
                        return ColorValue.Auto;
                    }
                    else
                    {
                        return new TableBorder(_doc, _style, TableRegionType.WholeTable, _type).Color;
                    }
                }
                return border.Color?.Value ?? ColorValue.Auto;
            }
            set
            {
                CreateBorder();
                if (_region == TableRegionType.WholeTable)
                {
                    if (_type == TableCellBorderType.Left)
                        _style.StyleTableProperties.TableBorders.LeftBorder.Color = value.ToString();
                    else if (_type == TableCellBorderType.Top)
                        _style.StyleTableProperties.TableBorders.TopBorder.Color = value.ToString();
                    else if (_type == TableCellBorderType.Right)
                        _style.StyleTableProperties.TableBorders.RightBorder.Color = value.ToString();
                    else if (_type == TableCellBorderType.Bottom)
                        _style.StyleTableProperties.TableBorders.BottomBorder.Color = value.ToString();
                    else if (_type == TableCellBorderType.InsideH)
                        _style.StyleTableProperties.TableBorders.InsideHorizontalBorder.Color = value.ToString();
                    else if (_type == TableCellBorderType.InsideV)
                        _style.StyleTableProperties.TableBorders.InsideVerticalBorder.Color = value.ToString();
                }
                else
                {
                    W.TableStyleOverrideValues type = _region.Convert<W.TableStyleOverrideValues>();
                    W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                    if (_type == TableCellBorderType.Left)
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.LeftBorder.Color = value.ToString();
                    else if (_type == TableCellBorderType.Top)
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.TopBorder.Color = value.ToString();
                    else if (_type == TableCellBorderType.Right)
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.RightBorder.Color = value.ToString();
                    else if (_type == TableCellBorderType.Bottom)
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.BottomBorder.Color = value.ToString();
                    else if (_type == TableCellBorderType.InsideH)
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideHorizontalBorder.Color = value.ToString();
                    else if (_type == TableCellBorderType.InsideV)
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideVerticalBorder.Color = value.ToString();
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
                TryGetBorder(out W.BorderType border);
                if (border == null)
                {
                    if (_region == TableRegionType.WholeTable)
                    {
                        W.Style baseStyle = _style.GetBaseStyle(_doc);
                        if (baseStyle != null)
                        {
                            return new TableBorder(_doc, baseStyle, _region, _type).Width;
                        }
                        return 0;
                    }
                    else
                    {
                        return new TableBorder(_doc, _style, TableRegionType.WholeTable, _type).Width;
                    }
                }
                if ((int)Style < 27)
                    return border.Size.Value / 8.0F;
                else
                    return border.Size.Value;
            }
            set
            {
                CreateBorder();
                W.BorderType border = null;
                if (_region == TableRegionType.WholeTable)
                {
                    if (_type == TableCellBorderType.Left)
                        border = _style.StyleTableProperties.TableBorders.LeftBorder;
                    else if (_type == TableCellBorderType.Top)
                        border = _style.StyleTableProperties.TableBorders.TopBorder;
                    else if (_type == TableCellBorderType.Right)
                        border = _style.StyleTableProperties.TableBorders.RightBorder;
                    else if (_type == TableCellBorderType.Bottom)
                        border = _style.StyleTableProperties.TableBorders.BottomBorder;
                    else if (_type == TableCellBorderType.InsideH)
                        border = _style.StyleTableProperties.TableBorders.InsideHorizontalBorder;
                    else if (_type == TableCellBorderType.InsideV)
                        border = _style.StyleTableProperties.TableBorders.InsideVerticalBorder;
                }
                else
                {
                    W.TableStyleOverrideValues type = _region.Convert<W.TableStyleOverrideValues>();
                    W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                    if (_type == TableCellBorderType.Left)
                        border = tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.LeftBorder;
                    else if (_type == TableCellBorderType.Top)
                        border = tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.TopBorder;
                    else if (_type == TableCellBorderType.Right)
                        border = tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.RightBorder;
                    else if (_type == TableCellBorderType.Bottom)
                        border = tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.BottomBorder;
                    else if (_type == TableCellBorderType.InsideH)
                        border = tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideHorizontalBorder;
                    else if (_type == TableCellBorderType.InsideV)
                        border = tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideVerticalBorder;
                }
                if(border != null)
                {
                    if ((int)Style < 27)
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
        }
        #endregion

        #region Private Methods
        private void TryGetBorder(out W.BorderType border)
        {
            border = null;
            if (_region == TableRegionType.WholeTable)
            {
                if(_type == TableCellBorderType.Left)
                    border = _style.StyleTableProperties?.TableBorders?.LeftBorder;
                else if (_type == TableCellBorderType.Top)
                    border = _style.StyleTableProperties?.TableBorders?.TopBorder;
                else if (_type == TableCellBorderType.Bottom)
                    border = _style.StyleTableProperties?.TableBorders?.BottomBorder;
                else if (_type == TableCellBorderType.Left)
                    border = _style.StyleTableProperties?.TableBorders?.LeftBorder;
                else if (_type == TableCellBorderType.Right)
                    border = _style.StyleTableProperties?.TableBorders?.RightBorder;
                else if (_type == TableCellBorderType.InsideH)
                    border = _style.StyleTableProperties?.TableBorders?.InsideHorizontalBorder;
                else if (_type == TableCellBorderType.InsideV)
                    border = _style.StyleTableProperties?.TableBorders?.InsideVerticalBorder;
            }
            else
            {
                W.TableStyleOverrideValues type = _region.Convert<W.TableStyleOverrideValues>();
                W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                if (_type == TableCellBorderType.Left)
                    border = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.LeftBorder;
                else if (_type == TableCellBorderType.Top)
                    border = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.TopBorder;
                else if (_type == TableCellBorderType.Bottom)
                    border = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.BottomBorder;
                else if (_type == TableCellBorderType.Left)
                    border = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.LeftBorder;
                else if (_type == TableCellBorderType.Right)
                    border = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.RightBorder;
                else if (_type == TableCellBorderType.InsideH)
                    border = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.InsideHorizontalBorder;
                else if (_type == TableCellBorderType.InsideV)
                    border = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.InsideVerticalBorder;
            }
        }

        private void CreateBorder()
        {
            if (_region == TableRegionType.WholeTable)
            {
                if(_style.StyleTableProperties == null)
                {
                    _style.StyleTableProperties = new W.StyleTableProperties();
                }
                if(_style.StyleTableProperties.TableBorders == null)
                {
                    _style.StyleTableProperties.TableBorders = new W.TableBorders();
                }

                if (_type == TableCellBorderType.Left
                    && _style.StyleTableProperties.TableBorders.LeftBorder == null)
                {
                    _style.StyleTableProperties.TableBorders.LeftBorder = new W.LeftBorder();
                }
                else if (_type == TableCellBorderType.Top
                    && _style.StyleTableProperties.TableBorders.TopBorder == null)
                {
                    _style.StyleTableProperties.TableBorders.TopBorder = new W.TopBorder();
                }
                else if (_type == TableCellBorderType.Bottom
                    && _style.StyleTableProperties.TableBorders.BottomBorder == null)
                {
                    _style.StyleTableProperties.TableBorders.BottomBorder = new W.BottomBorder();
                }
                else if (_type == TableCellBorderType.Right
                    && _style.StyleTableProperties.TableBorders.RightBorder == null)
                {
                    _style.StyleTableProperties.TableBorders.RightBorder = new W.RightBorder();
                }
                else if (_type == TableCellBorderType.InsideH
                    && _style.StyleTableProperties.TableBorders.InsideHorizontalBorder == null)
                {
                    _style.StyleTableProperties.TableBorders.InsideHorizontalBorder = new W.InsideHorizontalBorder();
                }
                else if (_type == TableCellBorderType.InsideV
                    && _style.StyleTableProperties.TableBorders.InsideVerticalBorder == null)
                {
                    _style.StyleTableProperties.TableBorders.InsideVerticalBorder = new W.InsideVerticalBorder();
                }
            }
            else
            {
                W.TableStyleOverrideValues type = _region.Convert<W.TableStyleOverrideValues>();
                if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                {
                    _style.Append(new W.TableStyleProperties() { Type = type });
                }
                W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                if (tblStylePr.TableStyleConditionalFormattingTableCellProperties == null)
                    tblStylePr.TableStyleConditionalFormattingTableCellProperties = new W.TableStyleConditionalFormattingTableCellProperties();
                if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders == null)
                    tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders = new W.TableCellBorders();

                if (_type == TableCellBorderType.Left
                    && tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.LeftBorder == null)
                {
                    tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.LeftBorder = new W.LeftBorder();
                }
                else if (_type == TableCellBorderType.Top
                    && tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.TopBorder == null)
                {
                    tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.TopBorder = new W.TopBorder();
                }
                else if (_type == TableCellBorderType.Bottom
                    && tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.BottomBorder == null)
                {
                    tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.BottomBorder = new W.BottomBorder();
                }
                else if (_type == TableCellBorderType.Right
                    && tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.RightBorder == null)
                {
                    tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.RightBorder = new W.RightBorder();
                }
                else if (_type == TableCellBorderType.InsideH
                    && tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideHorizontalBorder == null)
                {
                    tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideHorizontalBorder = new W.InsideHorizontalBorder();
                }
                else if (_type == TableCellBorderType.InsideV
                    && tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideVerticalBorder == null)
                {
                    tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.InsideVerticalBorder = new W.InsideVerticalBorder();
                }
            }
        }
        #endregion
    }



    internal enum TableCellBorderType
    {
        Left = 0,
        Top = 1,
        Right = 2,
        Bottom = 3,
        InsideH = 4,
        InsideV = 5
    }
}
