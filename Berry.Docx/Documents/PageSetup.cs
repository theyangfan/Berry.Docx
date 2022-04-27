using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Collections;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// Represent the page setup.
    /// </summary>
    public class PageSetup
    {
        #region Private Members
        private readonly Document _doc;
        private readonly Section _sect;
        private readonly W.PageSize _pgSz;
        private readonly W.PageMargin _pgMar;
        private W.VerticalTextAlignmentOnPage _vAlign;
        private W.DocGrid _docGrid;
        private W.TextDirection _textDirection;
        private W.Columns _columns;
        #endregion

        #region Constructors
        internal PageSetup(Document doc, Section section)
        {
            _doc = doc;
            _sect = section;
            _pgSz = section.XElement.GetFirstChild<W.PageSize>();
            _pgMar = section.XElement.GetFirstChild<W.PageMargin>();
            _vAlign = section.XElement.GetFirstChild<W.VerticalTextAlignmentOnPage>();
            _docGrid = section.XElement.GetFirstChild<W.DocGrid>();
            _textDirection = section.XElement.GetFirstChild<W.TextDirection>();
            _columns = section.XElement.GetFirstChild<W.Columns>();
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets or sets the size (in points) for all pages in the current section.
        /// </summary>
        public SizeF PageSize
        {
            get
            {
                return new SizeF(PageWidth, PageHeight);
            }
            set
            {
                PageWidth = value.Width;
                PageHeight = value.Height;
            }
        }
        /// <summary>
        /// Gets or sets the width (in points) for all pages in the current section.
        /// </summary>
        public float PageWidth
        {
            get
            {
                if (_pgSz?.Width == null) return 0;
                return (_pgSz.Width / 20.0F).Round(1);
            }
            set
            {
                _pgSz.Width = (uint)((value * 20).Round(0));
            }
        }
        /// <summary>
        /// Gets or sets the height (in points) for all pages in the current section.
        /// </summary>
        public float PageHeight
        {
            get
            {
                if (_pgSz?.Height == null) return 0;
                return (_pgSz.Height / 20.0F).Round(1);
            }
            set
            {
                _pgSz.Height = (uint)((value * 20).Round(0));
            }
        }

        /// <summary>
        /// Gets or sets the orientation of all pages in this section. The actual paper size width and height
        /// will be reversed for pages in this section if the orientation changed.
        /// </summary>
        public PageOrientation Orientation
        {
            get
            {
                if (_pgSz.Orient == null) return PageOrientation.Portrait;
                if (_pgSz.Orient == W.PageOrientationValues.Landscape) return PageOrientation.Landscape;
                return PageOrientation.Portrait;
            }
            set
            {
                if(value == PageOrientation.Portrait)
                {
                    if(_pgSz.Orient == W.PageOrientationValues.Landscape)
                    {
                        PageSize = new SizeF(PageHeight, PageWidth);
                    }
                    _pgSz.Orient = null;
                }
                else
                {
                    if (_pgSz.Orient == null || _pgSz.Orient == W.PageOrientationValues.Portrait)
                    {
                        PageSize = new SizeF(PageHeight, PageWidth);
                    }
                    _pgSz.Orient = W.PageOrientationValues.Landscape;
                }
            }
        }

        /// <summary>
        /// Gets or sets the page margins (in points) for all pages in this section.
        /// </summary>
        public MarginsF Margins
        {
            get
            {
                return new MarginsF(LeftMargin, RightMargin, TopMargin, BottomMargin);
            }
            set
            {
                LeftMargin = value.Left;
                RightMargin = value.Right;
                TopMargin = value.Top;
                BottomMargin = value.Bottom;
            }
        }

        /// <summary>
        /// Gets or sets the distance (in points) between the left edge of the text extents for this document and
        /// the left edge of the page for all pages in this section.
        /// </summary>
        public float LeftMargin
        {
            get
            {
                if (_pgMar?.Left == null) return 0;
                return (_pgMar.Left / 20.0F).Round(2);
            }
            set
            {
                _pgMar.Left = (uint)((value * 20).Round(0));
            }
        }

        /// <summary>
        /// Gets or sets the distance (in points) between the right edge of the text extents for this document and
        /// the right edge of the page for all pages in this section.
        /// </summary>
        public float RightMargin
        {
            get
            {
                if (_pgMar?.Right == null) return 0;
                return (_pgMar.Right / 20.0F).Round(2);
            }
            set
            {
                _pgMar.Right = (uint)((value * 20).Round(0));
            }
        }

        /// <summary>
        /// Gets or sets the distance (in points) between the top of the text margins for
        /// the main document and the top of the page for all pages in this section.
        /// </summary>
        public float TopMargin
        {
            get
            {
                if (_pgMar?.Top == null) return 0;
                return (_pgMar.Top / 20.0F).Round(2);
            }
            set
            {
                _pgMar.Top = (int)((value * 20).Round(0));
            }
        }

        /// <summary>
        /// Gets or sets the distance (in points) between the bottom of the text margins for
        /// the main document and the bottom of the page for all pages in this section.
        /// </summary>
        public float BottomMargin
        {
            get
            {
                if (_pgMar?.Bottom == null) return 0;
                return (_pgMar.Bottom / 20.0F).Round(2);
            }
            set
            {
                _pgMar.Bottom = (int)((value * 20).Round(0));
            }
        }

        /// <summary>
        /// Gets or sets the page gutter (in points) for each page in the current section.
        /// </summary>
        public float Gutter
        {
            get
            {
                if (_pgMar?.Gutter == null) return 0;
                return (_pgMar.Gutter / 20.0F).Round(2);
            }
            set
            {
                _pgMar.Gutter = (uint)((value * 20).Round(0));
            }
        }

        /// <summary>
        /// Gets or sets the page gutter location for each page in the current section.
        /// </summary>
        public GutterLocation GutterLocation
        {
            get
            {
                return _doc.Settings.GutterAtTop ? GutterLocation.Top : GutterLocation.Left;
            }
            set
            {
                if(value == GutterLocation.Left)
                    _doc.Settings.GutterAtTop = false;
                else
                    _doc.Settings.GutterAtTop = true;
            }
        }

        /// <summary>
        /// Gets or sets the distance (in points) from the top edge of the page to the top
        /// edge of the header.
        /// </summary>
        public float HeaderDistance
        {
            get
            {
                if (_pgMar?.Header == null) return 0;
                return (_pgMar.Header / 20.0F).Round(2);
            }
            set
            {
                _pgMar.Header = (uint)((value * 20).Round(0));
            }
        }

        /// <summary>
        /// Gets or sets the distance (in points) from the bottom edge of the page to the
        /// bottom edge of the footer.
        /// </summary>
        public float FooterDistance
        {
            get
            {
                if (_pgMar?.Footer == null) return 0;
                return (_pgMar.Footer / 20.0F).Round(2);
            }
            set
            {
                _pgMar.Footer = (uint)((value * 20).Round(0));
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public MultiPage MultiPage
        {
            get
            {
                if (_doc.Settings.MirrorMargins)
                    return MultiPage.MirrorMargins;
                else if (_doc.Settings.PrintTwoOnOne)
                    return MultiPage.PrintTwoOnOne;
                else
                    return MultiPage.Normal;
            }
            set
            {
                switch (value)
                {
                    case MultiPage.MirrorMargins:
                        _doc.Settings.MirrorMargins = true;
                        _doc.Settings.PrintTwoOnOne = false;
                        break;
                    case MultiPage.PrintTwoOnOne:
                        _doc.Settings.MirrorMargins = false;
                        _doc.Settings.PrintTwoOnOne = true;
                        break;
                    default:
                        _doc.Settings.MirrorMargins = false;
                        _doc.Settings.PrintTwoOnOne = false;
                        break;
                }
            }
        }

        public VerticalJustificationType VerticalJustification
        {
            get
            {
                if (_vAlign?.Val == null) return VerticalJustificationType.Top;
                return (VerticalJustificationType)(int)_vAlign.Val.Value;
            }
            set
            {
                if(_vAlign == null)
                {
                    _vAlign = new W.VerticalTextAlignmentOnPage();
                    _sect.XElement.AddChild(_vAlign);
                }
                _vAlign.Val = (W.VerticalJustificationValues)(int)value;
            }
        }

        public TextFlowDirection TextDirection
        {
            get
            {
                if (_textDirection?.Val == null) return TextFlowDirection.Horizontal;
                if (_textDirection.Val == W.TextDirectionValues.TopToBottomRightToLeft)
                    return TextFlowDirection.Vertical;
                else if (_textDirection.Val == W.TextDirectionValues.LefttoRightTopToBottomRotated)
                    return TextFlowDirection.RotateAsianChars270;
                else
                    return TextFlowDirection.Horizontal;
            }
            set
            {
                if(_textDirection == null)
                {
                    _textDirection = new W.TextDirection();
                    _sect.XElement.AddChild(_textDirection);
                }
                if (value == TextFlowDirection.Horizontal)
                    _textDirection.Val = W.TextDirectionValues.LefToRightTopToBottom;
                else if (value == TextFlowDirection.Vertical)
                    _textDirection.Val = W.TextDirectionValues.TopToBottomRightToLeft;
                else
                    _textDirection.Val = W.TextDirectionValues.LefttoRightTopToBottomRotated;
            }
        }

        public Columns Columns
        {
            get
            {
                if (_columns == null)
                {
                    _columns = new W.Columns();
                    _sect.XElement.AddChild(_columns);
                }
                return new Columns(_doc, _sect, _columns);
            }
        }

        public DocGridType DocGrid
        {
            get
            {
                if (_docGrid?.Type == null) return DocGridType.None;
                return (DocGridType)(int)_docGrid.Type.Value;
            }
            set
            {
                if (_docGrid == null)
                {
                    _docGrid = new W.DocGrid();
                    _sect.XElement.AddChild(_docGrid);
                }
                if (value == DocGridType.None) _docGrid.Type = null;
                else _docGrid.Type = (W.DocGridValues)(int)value;
            }
        }

        public float CharPitch
        {
            get
            {
                if (_docGrid?.CharacterSpace == null) return 0;
                ParagraphStyle normal = _doc.Styles.FindByName("normal", StyleType.Paragraph) as ParagraphStyle;
                float normalSz = normal?.CharacterFormat?.FontSize ?? 11.0F;
                return (_docGrid.CharacterSpace / 4096.0F + normalSz).Round(2);
            }
            set
            {
                if(_docGrid == null)
                {
                    _docGrid = new W.DocGrid();
                    _sect.XElement.AddChild(_docGrid);
                }
                ParagraphStyle normal = _doc.Styles.FindByName("normal", StyleType.Paragraph) as ParagraphStyle;
                float normalSz = normal?.CharacterFormat?.FontSize ?? 11.0F;
                _docGrid.CharacterSpace = (int)((value - normalSz) * 4096.0F).Round(0);
            }
        }

        public float LinePitch
        {
            get
            {
                if (_docGrid?.LinePitch == null) return 0;
                return (_docGrid.LinePitch / 20.0F).Round(2);
            }
            set
            {
                if (_docGrid == null)
                {
                    _docGrid = new W.DocGrid();
                    _sect.XElement.AddChild(_docGrid);
                }
                _docGrid.LinePitch = (int)(value * 20.0F).Round(0);
            }
        }
        #endregion
    }
}
