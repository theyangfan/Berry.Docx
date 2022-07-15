// Copyright (c) theyangfan. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Repersent the paragraph borders.
    /// </summary>
    public class Borders
    {

        #region Private Members
        private readonly Document _doc;
        private readonly W.Paragraph _ownerParagraph;
        private readonly W.Style _ownerStyle;
        #endregion

        #region Constructors
        internal Borders(Document doc, W.Paragraph paragraph)
        {
            _doc = doc;
            _ownerParagraph = paragraph;
        }

        internal Borders(Document doc, W.Style style)
        {
            _doc = doc;
            _ownerStyle = style;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The top border.
        /// </summary>
        public Border Top
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return new Border(_doc, _ownerParagraph, BorderType.Top);
                }
                else
                {
                    return new Border(_doc, _ownerStyle, BorderType.Top);
                }
            }
        }

        /// <summary>
        /// The bottom border.
        /// </summary>
        public Border Bottom
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return new Border(_doc, _ownerParagraph, BorderType.Bottom);
                }
                else
                {
                    return new Border(_doc, _ownerStyle, BorderType.Bottom);
                }
            }
        }

        /// <summary>
        /// The left border.
        /// </summary>
        public Border Left
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return new Border(_doc, _ownerParagraph, BorderType.Left);
                }
                else
                {
                    return new Border(_doc, _ownerStyle, BorderType.Left);
                }
            }
        }

        /// <summary>
        /// The right border.
        /// </summary>
        public Border Right
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return new Border(_doc, _ownerParagraph, BorderType.Right);
                }
                else
                {
                    return new Border(_doc, _ownerStyle, BorderType.Right);
                }
            }
        }

        /// <summary>
        /// Paragraph Border Between Identical Paragraphs.
        /// </summary>
        public Border Between
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return new Border(_doc, _ownerParagraph, BorderType.Between);
                }
                else
                {
                    return new Border(_doc, _ownerStyle, BorderType.Between);
                }
            }
        }

        /// <summary>
        /// Paragraph Border Between Facing Pages.
        /// </summary>
        public Border Bar
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return new Border(_doc, _ownerParagraph, BorderType.Bar);
                }
                else
                {
                    return new Border(_doc, _ownerStyle, BorderType.Bar);
                }
            }
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Sets borders.
        /// </summary>
        /// <param name="style"></param>
        /// <param name="color"></param>
        /// <param name="width"></param>
        public void SetBorders(BorderStyle style, ColorValue color, float width)
        {
            Top.Style = style;
            Top.Color = color;
            Top.Width = width;
            Bottom.Style = style;
            Bottom.Color = color;
            Bottom.Width = width;
            Left.Style = style;
            Left.Color = color;
            Left.Width = width;
            Right.Style = style;
            Right.Color = color;
            Right.Width = width;
        }

        /// <summary>
        /// Clears borders.
        /// </summary>
        public void Clear()
        {
            if(_ownerParagraph?.ParagraphProperties?.ParagraphBorders != null)
            {
                _ownerParagraph.ParagraphProperties.ParagraphBorders = null;
            }
            else if(_ownerStyle != null)
            {
                if (_ownerStyle.StyleParagraphProperties?.ParagraphBorders != null)
                    _ownerStyle.StyleParagraphProperties.ParagraphBorders = null;
                // clear borders in base style.
                W.Style baseStyle = _ownerStyle.GetBaseStyle(_doc);
                if (baseStyle != null)
                {
                    new Borders(_doc, baseStyle).Clear();
                }
            }
        }
        #endregion
    }

    /// <summary>
    /// Repersent the paragraph border.
    /// </summary>
    public class Border
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.Paragraph _ownerParagraph;
        private readonly W.Style _ownerStyle;
        private readonly W.RunPropertiesDefault _defaultRPr;
        private readonly EnumValue<BorderType> _borderType;
        #endregion

        #region Constructors
        internal Border(Document doc, W.Run run)
        {
            _doc = doc;
            _ownerRun = run;
        }

        internal Border(Document doc, W.Paragraph paragraph, EnumValue<BorderType> type = null)
        {
            _doc = doc;
            _ownerParagraph = paragraph;
            _borderType = type;
        }

        internal Border(Document doc, W.Style style, EnumValue<BorderType> type = null)
        {
            _doc = doc;
            _ownerStyle = style;
            _borderType = type;
        }

        internal Border(Document doc, W.RunPropertiesDefault defaultRPr)
        {
            _doc = doc;
            _defaultRPr = defaultRPr;
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
                if (_ownerRun != null)
                {
                    // direct formatting
                    if (BorderProperty != null)
                        return BorderProperty.Val?.Value.Convert<BorderStyle>() ?? BorderStyle.None;
                    // character style
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        W.BorderType border = GetStyleBorderRecursively(_doc, _ownerRun.GetStyle(_doc), _borderType);
                        if (border != null)
                            return border.Val?.Value.Convert<BorderStyle>() ?? BorderStyle.None;
                    }
                    // paragraph style
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        W.BorderType border = GetStyleBorderRecursively(_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc), _borderType);
                        if (border != null)
                            return border.Val?.Value.Convert<BorderStyle>() ?? BorderStyle.None;
                    }
                }
                else if (_ownerParagraph != null)
                {
                    // direct formatting
                    if (BorderProperty != null)
                        return BorderProperty.Val?.Value.Convert<BorderStyle>() ?? BorderStyle.None;
                    // paragraph style
                    W.BorderType border = GetStyleBorderRecursively(_doc, _ownerParagraph.GetStyle(_doc), _borderType);
                    if (border != null)
                        return border.Val?.Value.Convert<BorderStyle>() ?? BorderStyle.None;
                }
                else if (_ownerStyle != null)
                {
                    W.BorderType border = GetStyleBorderRecursively(_doc, _ownerStyle, _borderType);
                    if (border != null)
                        return border.Val?.Value.Convert<BorderStyle>() ?? BorderStyle.None;
                }
                else if (_defaultRPr != null)
                {
                    if (BorderProperty != null)
                        return BorderProperty.Val?.Value.Convert<BorderStyle>() ?? BorderStyle.None;
                }
                return BorderStyle.None;
            }
            set
            {
                InitProperty();
                if (BorderProperty != null)
                {
                    BorderProperty.Val = value.Convert<W.BorderValues>();
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
                if (_ownerRun != null)
                {
                    // direct formatting
                    if (BorderProperty != null)
                        return BorderProperty.Color?.Value ?? ColorValue.Auto;
                    // character style
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        W.BorderType border = GetStyleBorderRecursively(_doc, _ownerRun.GetStyle(_doc), _borderType);
                        if (border != null)
                            return border.Color?.Value ?? ColorValue.Auto;
                    }
                    // paragraph style
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        W.BorderType border = GetStyleBorderRecursively(_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc), _borderType);
                        if (border != null)
                            return border.Color?.Value ?? ColorValue.Auto;
                    }
                }
                else if (_ownerParagraph != null)
                {
                    // direct formatting
                    if (BorderProperty != null)
                        return BorderProperty.Color?.Value ?? ColorValue.Auto;
                    // paragraph style
                    W.BorderType border = GetStyleBorderRecursively(_doc, _ownerParagraph.GetStyle(_doc), _borderType);
                    if (border != null)
                        return border.Color?.Value ?? ColorValue.Auto;
                }
                else if (_ownerStyle != null)
                {
                    W.BorderType border = GetStyleBorderRecursively(_doc, _ownerStyle, _borderType);
                    if (border != null)
                        return border.Color?.Value ?? ColorValue.Auto;
                }
                else if (_defaultRPr != null)
                {
                    if (BorderProperty != null)
                        return BorderProperty.Color?.Value ?? ColorValue.Auto;
                }
                return ColorValue.Auto;
            }
            set
            {
                InitProperty();
                if (BorderProperty != null)
                {
                    BorderProperty.Color = value.ToString();
                }
            }
        }

        /// <summary>
        /// Gets or sets the border width.
        /// </summary>
        public float Width
        {
            get
            {
                if (_ownerRun != null)
                {
                    // direct formatting
                    if (BorderProperty != null)
                    {
                        if (BorderProperty.Size == null) return 0;
                        if ((int)Style < 27)
                            return BorderProperty.Size.Value / 8.0F;
                        else
                            return BorderProperty.Size.Value;
                    }
                    // character style
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        W.BorderType border = GetStyleBorderRecursively(_doc, _ownerRun.GetStyle(_doc), _borderType);
                        if (border != null)
                        {
                            if (border.Size == null) return 0;
                            if ((int)Style < 27)
                                return border.Size.Value / 8.0F;
                            else
                                return border.Size.Value;
                        }
                    }
                    // paragraph style
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        W.BorderType border = GetStyleBorderRecursively(_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc), _borderType);
                        if (border != null)
                        {
                            if (border.Size == null) return 0;
                            if ((int)Style < 27)
                                return border.Size.Value / 8.0F;
                            else
                                return border.Size.Value;
                        }
                    }
                }
                else if (_ownerParagraph != null)
                {
                    // direct formatting
                    if (BorderProperty != null)
                    {
                        if (BorderProperty.Size == null) return 0;
                        if ((int)Style < 27)
                            return BorderProperty.Size.Value / 8.0F;
                        else
                            return BorderProperty.Size.Value;
                    }
                    // paragraph style
                    W.BorderType border = GetStyleBorderRecursively(_doc, _ownerParagraph.GetStyle(_doc), _borderType);
                    if (border != null)
                    {
                        if (border.Size == null) return 0;
                        if ((int)Style < 27)
                            return border.Size.Value / 8.0F;
                        else
                            return border.Size.Value;
                    }
                }
                else if (_ownerStyle != null)
                {
                    W.BorderType border = GetStyleBorderRecursively(_doc, _ownerStyle, _borderType);
                    if (border != null)
                    {
                        if (border.Size == null) return 0;
                        if ((int)Style < 27)
                            return border.Size.Value / 8.0F;
                        else
                            return border.Size.Value;
                    }
                }
                else if (_defaultRPr != null)
                {
                    if (BorderProperty != null)
                    {
                        if (BorderProperty.Size == null) return 0;
                        if ((int)Style < 27)
                            return BorderProperty.Size.Value / 8.0F;
                        else
                            return BorderProperty.Size.Value;
                    }
                }
                return 0;
            }
            set
            {
                InitProperty();
                if (BorderProperty != null)
                {
                    if ((int)Style < 27)
                    {
                        if (value > 12)
                            BorderProperty.Size = 96;
                        else if (value >= 0.25)
                            BorderProperty.Size = (uint)(value * 8);
                        else if (value > 0)
                            BorderProperty.Size = 2;
                        else
                            BorderProperty.Size = 0;
                    }
                    else
                    {
                        if (value > 31)
                            BorderProperty.Size = 31;
                        else if (value >= 1)
                            BorderProperty.Size = (uint)value;
                        else if (value > 0)
                            BorderProperty.Size = 1;
                        else
                            BorderProperty.Size = 0;
                    }
                }
            }
        }
        #endregion

        #region Internal Properties
        internal W.BorderType BorderProperty
        {
            get
            {
                if (_ownerRun != null)
                {
                    return _ownerRun.RunProperties?.Border;
                }
                else if (_ownerParagraph != null)
                {
                    if (_borderType == null)
                        return _ownerParagraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.Border>();
                    if (_borderType == BorderType.Top)
                        return _ownerParagraph.ParagraphProperties?.ParagraphBorders?.TopBorder;
                    else if (_borderType == BorderType.Bottom)
                        return _ownerParagraph.ParagraphProperties?.ParagraphBorders?.BottomBorder;
                    else if (_borderType == BorderType.Left)
                        return _ownerParagraph.ParagraphProperties?.ParagraphBorders?.LeftBorder;
                    else if (_borderType == BorderType.Right)
                        return _ownerParagraph.ParagraphProperties?.ParagraphBorders?.RightBorder;
                    else if (_borderType == BorderType.Between)
                        return _ownerParagraph.ParagraphProperties?.ParagraphBorders?.BetweenBorder;
                    else if (_borderType == BorderType.Bar)
                        return _ownerParagraph.ParagraphProperties?.ParagraphBorders?.BarBorder;
                }
                else if (_ownerStyle != null)
                {
                    if (_borderType == null)
                        return _ownerStyle.StyleRunProperties?.Border;
                    if (_borderType == BorderType.Top)
                        return _ownerStyle.StyleParagraphProperties?.ParagraphBorders?.TopBorder;
                    else if (_borderType == BorderType.Bottom)
                        return _ownerStyle.StyleParagraphProperties?.ParagraphBorders?.BottomBorder;
                    else if (_borderType == BorderType.Left)
                        return _ownerStyle.StyleParagraphProperties?.ParagraphBorders?.LeftBorder;
                    else if (_borderType == BorderType.Right)
                        return _ownerStyle.StyleParagraphProperties?.ParagraphBorders?.RightBorder;
                    else if (_borderType == BorderType.Between)
                        return _ownerStyle.StyleParagraphProperties?.ParagraphBorders?.BetweenBorder;
                    else if (_borderType == BorderType.Bar)
                        return _ownerStyle.StyleParagraphProperties?.ParagraphBorders?.BarBorder;
                }
                else if (_defaultRPr != null)
                {
                    return _defaultRPr.RunPropertiesBaseStyle?.Border;
                }
                return null;
            }
        }
        #endregion

        #region Private Methods
        private void InitProperty()
        {
            if (_ownerRun != null)
            {
                if (_ownerRun.RunProperties == null)
                    _ownerRun.RunProperties = new W.RunProperties();
                if (_ownerRun.RunProperties.Border == null)
                    _ownerRun.RunProperties.Border = new W.Border() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 };
            }
            else if (_ownerParagraph != null)
            {
                if (_ownerParagraph.ParagraphProperties == null)
                    _ownerParagraph.ParagraphProperties = new W.ParagraphProperties();
                if (_borderType != null)
                {
                    if (_ownerParagraph.ParagraphProperties.ParagraphBorders == null)
                        _ownerParagraph.ParagraphProperties.ParagraphBorders = new W.ParagraphBorders();
                    W.ParagraphBorders pBdr = _ownerParagraph.ParagraphProperties.ParagraphBorders;
                    if (_borderType == BorderType.Left && pBdr.LeftBorder == null)
                        pBdr.LeftBorder = new W.LeftBorder() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 };
                    else if (_borderType == BorderType.Top && pBdr.TopBorder == null)
                        pBdr.TopBorder = new W.TopBorder() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 };
                    else if (_borderType == BorderType.Right && pBdr.RightBorder == null)
                        pBdr.RightBorder = new W.RightBorder() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 };
                    else if (_borderType == BorderType.Bottom && pBdr.BottomBorder == null)
                        pBdr.BottomBorder = new W.BottomBorder() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 };
                    else if (_borderType == BorderType.Between && pBdr.BetweenBorder == null)
                        pBdr.BetweenBorder = new W.BetweenBorder() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 };
                    else if (_borderType == BorderType.Bar && pBdr.BarBorder == null)
                        pBdr.BarBorder = new W.BarBorder() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 };
                }
                else
                {
                    if (_ownerParagraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _ownerParagraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (!_ownerParagraph.ParagraphProperties.ParagraphMarkRunProperties.Elements<W.Border>().Any())
                        _ownerParagraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.Border() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 });
                }

            }
            else if (_ownerStyle != null)
            {
                if (_borderType != null)
                {
                    if (_ownerStyle.StyleParagraphProperties == null)
                        _ownerStyle.StyleParagraphProperties = new W.StyleParagraphProperties();
                    if (_ownerStyle.StyleParagraphProperties.ParagraphBorders == null)
                        _ownerStyle.StyleParagraphProperties.ParagraphBorders = new W.ParagraphBorders();
                    W.ParagraphBorders pBdr = _ownerStyle.StyleParagraphProperties.ParagraphBorders;
                    if (_borderType == BorderType.Left && pBdr.LeftBorder == null)
                        pBdr.LeftBorder = new W.LeftBorder() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 };
                    else if (_borderType == BorderType.Top && pBdr.TopBorder == null)
                        pBdr.TopBorder = new W.TopBorder() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 };
                    else if (_borderType == BorderType.Right && pBdr.RightBorder == null)
                        pBdr.RightBorder = new W.RightBorder() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 };
                    else if (_borderType == BorderType.Bottom && pBdr.BottomBorder == null)
                        pBdr.BottomBorder = new W.BottomBorder() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 };
                    else if (_borderType == BorderType.Between && pBdr.BetweenBorder == null)
                        pBdr.BetweenBorder = new W.BetweenBorder() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 };
                    else if (_borderType == BorderType.Bar && pBdr.BarBorder == null)
                        pBdr.BarBorder = new W.BarBorder() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 };
                }
                else
                {
                    if (_ownerStyle.StyleRunProperties == null)
                        _ownerStyle.StyleRunProperties = new W.StyleRunProperties();
                    if (_ownerStyle.StyleRunProperties.Border == null)
                        _ownerStyle.StyleRunProperties.Border = new W.Border() { Val = W.BorderValues.None, Color = "auto", Size = 0, Space = 0 };
                }
            }
        }

        private static W.BorderType GetStyleBorderRecursively(Document doc, W.Style style, EnumValue<BorderType> type)
        {
            W.BorderType baseBdr = null;
            W.Style baseStyle = style.GetBaseStyle(doc);
            if (baseStyle != null)
            {
                baseBdr = GetStyleBorderRecursively(doc, baseStyle, type);
            }
            if (type != null)
            {
                if (type == BorderType.Top)
                    return style.StyleParagraphProperties?.ParagraphBorders?.TopBorder ?? baseBdr;
                else if (type == BorderType.Bottom)
                    return style.StyleParagraphProperties?.ParagraphBorders?.BottomBorder ?? baseBdr;
                else if (type == BorderType.Left)
                    return style.StyleParagraphProperties?.ParagraphBorders?.LeftBorder ?? baseBdr;
                else if (type == BorderType.Right)
                    return style.StyleParagraphProperties?.ParagraphBorders?.RightBorder ?? baseBdr;
                else if (type == BorderType.Between)
                    return style.StyleParagraphProperties?.ParagraphBorders?.BetweenBorder ?? baseBdr;
                else
                    return style.StyleParagraphProperties?.ParagraphBorders?.BarBorder ?? baseBdr;
            }
            else
            {
                return style.StyleRunProperties?.Border ?? baseBdr;
            }
        }
        #endregion
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
