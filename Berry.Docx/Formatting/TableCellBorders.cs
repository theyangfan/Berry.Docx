﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    public class TableCellBorders
    {
        private readonly TableCellBorder _left;
        private readonly TableCellBorder _top;
        private readonly TableCellBorder _right;
        private readonly TableCellBorder _bottom;
        private readonly TableCellBorder _insideH;
        private readonly TableCellBorder _insideV;

        internal TableCellBorders(Document doc, W.Style style, TableRegionType region)
        {
            _left = new TableCellBorder(doc, style, region, TableCellBorderType.Left);
            _right = new TableCellBorder(doc, style, region, TableCellBorderType.Right);
            _top = new TableCellBorder(doc, style, region, TableCellBorderType.Top);
            _bottom = new TableCellBorder(doc, style, region, TableCellBorderType.Bottom);
            _insideH = new TableCellBorder(doc, style, region, TableCellBorderType.InsideH);
            _insideV = new TableCellBorder(doc, style, region, TableCellBorderType.InsideV);
        }

        public TableCellBorder Top => _top;
        public TableCellBorder Bottom => _bottom;
        public TableCellBorder Left => _left;
        public TableCellBorder Right => _right;
        public TableCellBorder InsideH => _insideH;
        public TableCellBorder InsideV => _insideV;
    }

    public class TableCellBorder
    {
        private readonly Document _doc;
        private readonly W.Style _style;
        private readonly TableRegionType _region;
        private readonly TableCellBorderType _type;

        internal TableCellBorder(Document doc, W.Style style, TableRegionType region, TableCellBorderType type)
        {
            _doc = doc;
            _style = style;
            _region = region;
            _type = type;
        }

        public BorderStyle Style
        {
            get
            {
                return BorderStyle.None;
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

        public ColorValue Color
        {
            get
            {
                return ColorValue.Auto;
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

        public float Width
        {
            get
            {
                return 0;
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

        private void TryGetBorder(out W.BorderType border)
        {
            border = null;
            if (_region == TableRegionType.WholeTable)
            {
                if(_type == TableCellBorderType.Left)
                    border = _style.StyleTableProperties?.TableBorders?.LeftBorder;
                else if (_type == TableCellBorderType.Top)
                    border = _style.StyleTableProperties?.TableBorders?.TopBorder;
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