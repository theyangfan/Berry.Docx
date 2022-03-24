using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// Represent the page setup.
    /// </summary>
    public class PageSetup
    {
        #region Private Members
        private readonly W.PageSize _pgSz;
        private readonly W.PageMargin _pgMar;
        private W.DocGrid _docGrid;
        #endregion

        #region Constructors
        internal PageSetup(Document doc, Section section)
        {
            _pgSz = section.XElement.GetFirstChild<W.PageSize>();
            _pgMar = section.XElement.GetFirstChild<W.PageMargin>();
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

        public float LeftMargin
        {
            get
            {
                return (_pgMar.Left / 20.0F).Round(2);
            }
            set
            {
                _pgMar.Left = (uint)((value * 20).Round(0));
            }
        }

        public float RightMargin
        {
            get
            {
                return (_pgMar.Right / 20.0F).Round(2);
            }
            set
            {
                _pgMar.Right = (uint)((value * 20).Round(0));
            }
        }

        public float TopMargin
        {
            get
            {
                return (_pgMar.Top / 20.0F).Round(2);
            }
            set
            {
                _pgMar.Top = (int)((value * 20).Round(0));
            }
        }

        public float BottomMargin
        {
            get
            {
                return (_pgMar.Bottom / 20.0F).Round(2);
            }
            set
            {
                _pgMar.Bottom = (int)((value * 20).Round(0));
            }
        }
        #endregion

        #region TODO
        private DocGridType DocGridType
        {
            get
            {
                if (_docGrid.Type == null) return DocGridType.None;
                if (_docGrid.Type == W.DocGridValues.Lines)
                    return DocGridType.Lines;
                else if (_docGrid.Type == W.DocGridValues.LinesAndChars)
                    return DocGridType.LinesAndChars;
                else if (_docGrid.Type == W.DocGridValues.SnapToChars)
                    return DocGridType.SnapToChars;
                else
                    return DocGridType.None;
            }
            set
            {
                if (value == DocGridType.Lines)
                    _docGrid.Type = W.DocGridValues.Lines;
                else if (value == DocGridType.LinesAndChars)
                    _docGrid.Type = W.DocGridValues.LinesAndChars;
                else if (value == DocGridType.SnapToChars)
                    _docGrid.Type = W.DocGridValues.SnapToChars;
                else
                    _docGrid.Type = null;
            }
        }
        #endregion
    }
}
