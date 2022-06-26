using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    public class Border
    {
        private readonly W.Run _run;
        private readonly W.Paragraph _paragraph;
        private readonly W.Style _style;
        private readonly W.RunPropertiesDefault _defaultRPr;
        private readonly EnumValue<BorderType> _borderType;

        internal Border(W.Run run)
        {
            _run = run;
        }

        internal Border(W.Paragraph paragraph, EnumValue<BorderType> type = null)
        {
            _paragraph = paragraph;
            _borderType = type;
        }

        internal Border(W.Style style, EnumValue<BorderType> type = null)
        {
            _style = style;
            _borderType = type;
        }

        internal Border(W.RunPropertiesDefault defaultRPr)
        {
            _defaultRPr = defaultRPr;
        }

        internal bool IsNUll => BorderProperty == null;

        internal W.BorderType BorderProperty
        {
            get
            {
                if (_run != null)
                {
                    return _run.RunProperties?.Border;
                }
                else if (_paragraph != null)
                {
                    if (_borderType == null)
                        return _paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.Border>();
                    if (_borderType == BorderType.Top)
                        return _paragraph.ParagraphProperties?.ParagraphBorders?.TopBorder;
                    else if (_borderType == BorderType.Bottom)
                        return _paragraph.ParagraphProperties?.ParagraphBorders?.BottomBorder;
                    else if (_borderType == BorderType.Left)
                        return _paragraph.ParagraphProperties?.ParagraphBorders?.LeftBorder;
                    else
                        return _paragraph.ParagraphProperties?.ParagraphBorders?.RightBorder;
                }
                else if(_style != null)
                {
                    if (_borderType == null)
                        return _style.StyleRunProperties?.Border;
                    if (_borderType == BorderType.Top)
                        return _style.StyleParagraphProperties?.ParagraphBorders?.TopBorder;
                    else if (_borderType == BorderType.Bottom)
                        return _style.StyleParagraphProperties?.ParagraphBorders?.BottomBorder;
                    else if (_borderType == BorderType.Left)
                        return _style.StyleParagraphProperties?.ParagraphBorders?.LeftBorder;
                    else
                        return _style.StyleParagraphProperties?.ParagraphBorders?.RightBorder;
                }
                else if(_defaultRPr != null)
                {
                    return _defaultRPr.RunPropertiesBaseStyle?.Border;
                }
                return null;
            }
        }

        public BorderStyle Style
        {
            get
            {
                return BorderProperty?.Val?.Value.Convert<BorderStyle>() ?? BorderStyle.None;
            }
            set
            {
                InitProperty();
                if(BorderProperty != null)
                {
                    BorderProperty.Val = value.Convert<W.BorderValues>();
                }
            }
        }

        private void InitProperty()
        {
            if(_run != null)
            {
                if(_run.RunProperties == null) 
                    _run.RunProperties = new W.RunProperties();
                if(_run.RunProperties.Border == null)
                    _run.RunProperties.Border = new W.Border();
            }
            else if(_paragraph != null)
            {
                if(_paragraph.ParagraphProperties == null)
                    _paragraph.ParagraphProperties = new W.ParagraphProperties();
                if(_borderType != null)
                {
                    if (_paragraph.ParagraphProperties.ParagraphBorders == null)
                        _paragraph.ParagraphProperties.ParagraphBorders = new W.ParagraphBorders();
                    W.ParagraphBorders pBdr = _paragraph.ParagraphProperties.ParagraphBorders;
                    if (_borderType == BorderType.Left && pBdr.LeftBorder == null)
                        pBdr.LeftBorder = new W.LeftBorder();
                    else if (_borderType == BorderType.Top && pBdr.TopBorder == null)
                        pBdr.TopBorder = new W.TopBorder();
                    else if (_borderType == BorderType.Right && pBdr.RightBorder == null)
                        pBdr.RightBorder = new W.RightBorder();
                    else if (_borderType == BorderType.Bottom && pBdr.BottomBorder == null)
                        pBdr.BottomBorder = new W.BottomBorder();
                }
                else
                {
                    if(_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if(!_paragraph.ParagraphProperties.ParagraphMarkRunProperties.Elements<W.Border>().Any())
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.Border());
                }
                
            }
            else if(_style != null)
            {
                if (_borderType != null)
                {
                    if (_style.StyleParagraphProperties == null)
                        _style.StyleParagraphProperties = new W.StyleParagraphProperties();
                    if (_style.StyleParagraphProperties.ParagraphBorders == null)
                        _style.StyleParagraphProperties.ParagraphBorders = new W.ParagraphBorders();
                    W.ParagraphBorders pBdr = _style.StyleParagraphProperties.ParagraphBorders;
                    if (_borderType == BorderType.Left && pBdr.LeftBorder == null)
                        pBdr.LeftBorder = new W.LeftBorder();
                    else if (_borderType == BorderType.Top && pBdr.TopBorder == null)
                        pBdr.TopBorder = new W.TopBorder();
                    else if (_borderType == BorderType.Right && pBdr.RightBorder == null)
                        pBdr.RightBorder = new W.RightBorder();
                    else if (_borderType == BorderType.Bottom && pBdr.BottomBorder == null)
                        pBdr.BottomBorder = new W.BottomBorder();
                }
                else
                {
                    if(_style.StyleRunProperties == null)
                        _style.StyleRunProperties = new W.StyleRunProperties();
                    if(_style.StyleRunProperties.Border == null)
                        _style.StyleRunProperties.Border = new W.Border();
                }
            }
        }
    }

    internal enum BorderType
    {
        Left = 0,
        Top = 1,
        Right = 2,
        Bottom = 3,
        Between = 4,
        Bar = 5
    }
}
