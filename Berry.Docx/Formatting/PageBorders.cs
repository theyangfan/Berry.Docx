using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represents the page borders.
    /// </summary>
    public class PageBorders
    {
        #region Private Members
        private readonly W.SectionProperties _sectPr;
        private readonly PageBorder _top;
        private readonly PageBorder _bottom;
        private readonly PageBorder _left;
        private readonly PageBorder _right;
        #endregion

        #region Constructors
        internal PageBorders(Document doc, Section section)
        {
            _sectPr = section.XElement;
            _top = new PageBorder(doc, section, PageBorderType.Top);
            _bottom = new PageBorder(doc, section, PageBorderType.Bottom);
            _left = new PageBorder(doc, section, PageBorderType.Left);
            _right = new PageBorder(doc, section, PageBorderType.Right);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets or sets the relative positioning of the page borders shall be calculated.
        /// </summary>
        public PageBordersPosition Position
        {
            get
            {
                W.PageBorders pgBorders = _sectPr.GetFirstChild<W.PageBorders>();
                return pgBorders?.OffsetFrom.Value.Convert<PageBordersPosition>() ?? PageBordersPosition.Text;
            }
            set
            {
                W.PageBorders pgBorders = _sectPr.GetFirstChild<W.PageBorders>();
                if (pgBorders == null)
                {
                    pgBorders = new W.PageBorders();
                    _sectPr.AddChild(pgBorders);
                }
                pgBorders.OffsetFrom = value.Convert<W.PageBorderOffsetValues>();
            }
        }

        /// <summary>
        /// Gets the top page border.
        /// </summary>
        public PageBorder Top => _top;

        /// <summary>
        /// Gets the bottom page border.
        /// </summary>
        public PageBorder Bottom => _bottom;

        /// <summary>
        /// Gets the left page border.
        /// </summary>
        public PageBorder Left => _left;

        /// <summary>
        /// Gets the right page border.
        /// </summary>
        public PageBorder Right => _right;
        #endregion

        #region Public Methods
        /// <summary>
        /// Clears all borders.
        /// </summary>
        public void Clear()
        {
            W.PageBorders pgBorders = _sectPr.GetFirstChild<W.PageBorders>();
            pgBorders?.Remove();
        }
        #endregion
    }

    /// <summary>
    /// Represents the page border.
    /// </summary>
    public class PageBorder
    {
        #region Private Members
        private readonly W.SectionProperties _sectPr;
        private readonly PageBorderType _type;
        #endregion

        #region Constructors
        internal PageBorder(Document doc, Section section, PageBorderType type)
        {
            _sectPr = section.XElement;
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
                if (border?.Val == null) return BorderStyle.Nil;
                return border.Val.Value.Convert<BorderStyle>();
            }
            set
            {
                CreateBorder();
                TryGetBorder(out W.BorderType border);
                border.Val = value.Convert<W.BorderValues>();
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
                if (border?.Color == null) return ColorValue.Auto;
                return border.Color.Value;
            }
            set
            {
                CreateBorder();
                TryGetBorder(out W.BorderType border);
                border.Color = value.ToString();
            }
        }

        /// <summary>
        /// Gets or sets the border width.
        /// </summary>
        public float Width
        {
            get
            {
                TryGetBorder(out W.BorderType border);
                if (border?.Size == null) return 0;
                if ((int)Style < 27)
                    return border.Size.Value / 8.0F;
                else
                    return border.Size.Value;
            }
            set
            {
                CreateBorder();
                TryGetBorder(out W.BorderType border);
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
        #endregion

        #region Private Methods
        private void TryGetBorder(out W.BorderType border)
        {
            border = null;
            W.PageBorders pgBorders = _sectPr.GetFirstChild<W.PageBorders>();
            if (pgBorders != null)
            {
                if (_type == PageBorderType.Top)
                    border = pgBorders.TopBorder;
                else if (_type == PageBorderType.Bottom)
                    border = pgBorders.BottomBorder;
                else if (_type == PageBorderType.Left)
                    border = pgBorders.LeftBorder;
                else if (_type == PageBorderType.Right)
                    border = pgBorders.RightBorder;
            }
        }

        private void CreateBorder()
        {
            W.PageBorders pgBorders = _sectPr.GetFirstChild<W.PageBorders>();
            if (pgBorders == null)
            {
                pgBorders = new W.PageBorders();
                _sectPr.AddChild(pgBorders);
            }
            if (_type == PageBorderType.Top && pgBorders.TopBorder == null)
            {
                pgBorders.TopBorder = new W.TopBorder() { Val = W.BorderValues.Nil };
            }
            else if (_type == PageBorderType.Bottom && pgBorders.BottomBorder == null)
            {
                pgBorders.BottomBorder = new W.BottomBorder() { Val = W.BorderValues.Nil };
            }
            else if (_type == PageBorderType.Left && pgBorders.LeftBorder == null)
            {
                pgBorders.LeftBorder = new W.LeftBorder() { Val = W.BorderValues.Nil };
            }
            else if (_type == PageBorderType.Right && pgBorders.RightBorder == null)
            {
                pgBorders.RightBorder = new W.RightBorder() { Val = W.BorderValues.Nil };
            }
        }
        #endregion

    }

    internal enum PageBorderType
    {
        Top = 0,
        Bottom = 1,
        Left = 2,
        Right = 3
    }
}
