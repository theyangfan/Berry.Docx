using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using P = DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx;
using Berry.Docx.Utils;
namespace Berry.Docx.Documents
{
    /// <summary>
    /// Represent the headers and footers of the section.
    /// </summary>
    public class HeaderFooters
    {
        #region Private Members
        private readonly Document _doc;
        private readonly Section _section;
        #endregion

        #region Constructors
        internal HeaderFooters(Document doc, Section section)
        {
            _doc = doc;
            _section = section;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets or sets a value indicating whether sections in this document shall have different headers and footers for even and
        /// odd pages (<see cref="OddHeader"/>/<see cref="OddFooter"/> and <see cref="EvenHeader"/>/<see cref="EvenFooter"/>). 
        /// This property will affect all sections of the document.
        /// </summary>
        public bool DifferentEvenAndOddHeaders
        {
            get => _doc.Settings.EvenAndOddHeaders;
            set => _doc.Settings.EvenAndOddHeaders = value;
        }

        /// <summary>
        /// Gets or sets a value indicating whether the parent section in this document shall have a different header and footer for
        /// its first page (see <see cref="FirstPageHeader"/> and <see cref="FirstPageFooter"/>).
        /// </summary>
        public bool DifferentFirstPageHeaders
        {
            get
            {
                return _section.XElement.Elements<W.TitlePage>().Any();
            }
            set
            {
                if (value)
                {
                    if (!_section.XElement.Elements<W.TitlePage>().Any())
                        _section.XElement.AddChild(new W.TitlePage());
                }
                else
                {
                    _section.XElement.RemoveAllChildren<W.TitlePage>();
                }
            }
        }

        /// <summary>
        /// Gets the default header of the section. This is equivalent to <see cref="OddHeader"/>.
        /// </summary>
        public HeaderFooter Header => OddHeader;

        /// <summary>
        /// Gets the first page header of the section.
        /// </summary>
        public HeaderFooter FirstPageHeader
        {
            get
            {
                W.HeaderReference first = _section.XElement.Elements<W.HeaderReference>().Where(h => h.Type == W.HeaderFooterValues.First).FirstOrDefault();
                if(first != null)
                {
                    P.HeaderPart headerPart = (P.HeaderPart)_doc.Package.MainDocumentPart.GetPartById(first.Id);
                    return new HeaderFooter(_doc, _section, headerPart.Header, first);
                }
                else
                {
                    if (_section.XElement.Elements<W.TitlePage>().Any()
                        && _section.PreviousSection?.HeaderFooters.FirstPageHeader != null)
                    {
                        return new HeaderFooter(_doc, _section, _section.PreviousSection.HeaderFooters.FirstPageHeader);
                    }
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets the odd page header of the section. This is also the default header.
        /// </summary>
        public HeaderFooter OddHeader
        {
            get
            {
                W.HeaderReference odd = _section.XElement.Elements<W.HeaderReference>().Where(h => h.Type == W.HeaderFooterValues.Default).FirstOrDefault();
                if (odd != null)
                {
                    P.HeaderPart headerPart = (P.HeaderPart)_doc.Package.MainDocumentPart.GetPartById(odd.Id);
                    return new HeaderFooter(_doc, _section, headerPart.Header, odd);
                }
                else
                {
                    if(_section.PreviousSection?.HeaderFooters.OddHeader != null)
                        return new HeaderFooter(_doc, _section, _section.PreviousSection.HeaderFooters.OddHeader);
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets the even page header of the section.
        /// </summary>
        public HeaderFooter EvenHeader
        {
            get
            {
                W.HeaderReference even = _section.XElement.Elements<W.HeaderReference>().Where(h => h.Type == W.HeaderFooterValues.Even).FirstOrDefault();
                if (even != null)
                {
                    P.HeaderPart headerPart = (P.HeaderPart)_doc.Package.MainDocumentPart.GetPartById(even.Id);
                    return new HeaderFooter(_doc, _section, headerPart.Header, even);
                }
                else
                {
                    if (_doc.Settings.EvenAndOddHeaders
                        && _section.PreviousSection?.HeaderFooters.EvenHeader != null)
                    {
                        return new HeaderFooter(_doc, _section, _section.PreviousSection.HeaderFooters.EvenHeader);
                    }
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets the default footer of the section. This is equivalent to <see cref="OddFooter"/>.
        /// </summary>
        public HeaderFooter Footer => OddFooter;

        /// <summary>
        /// Gets the first page footer of the section.
        /// </summary>
        public HeaderFooter FirstPageFooter
        {
            get
            {
                W.FooterReference first = _section.XElement.Elements<W.FooterReference>().Where(f => f.Type == W.HeaderFooterValues.First).FirstOrDefault();
                if (first != null)
                {
                    P.FooterPart footerPart = (P.FooterPart)_doc.Package.MainDocumentPart.GetPartById(first.Id);
                    return new HeaderFooter(_doc, _section, footerPart.Footer, first);
                }
                else
                {
                    if (_section.XElement.Elements<W.TitlePage>().Any()
                        && _section.PreviousSection?.HeaderFooters.FirstPageFooter != null)
                    {
                        return new HeaderFooter(_doc, _section, _section.PreviousSection.HeaderFooters.FirstPageFooter);
                    }
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets the odd page footer of the section. This is also the default footer.
        /// </summary>
        public HeaderFooter OddFooter
        {
            get
            {
                W.FooterReference odd = _section.XElement.Elements<W.FooterReference>().Where(f => f.Type == W.HeaderFooterValues.Default).FirstOrDefault();
                if (odd != null)
                {
                    P.FooterPart footerPart = (P.FooterPart)_doc.Package.MainDocumentPart.GetPartById(odd.Id);
                    return new HeaderFooter(_doc, _section, footerPart.Footer, odd);
                }
                else
                {
                    if (_section.PreviousSection?.HeaderFooters.OddFooter != null)
                        return new HeaderFooter(_doc, _section, _section.PreviousSection.HeaderFooters.OddFooter);
                    return null;
                }
            }
        }

        /// <summary>
        /// Gets the even page footer of the section.
        /// </summary>
        public HeaderFooter EvenFooter
        {
            get
            {
                W.FooterReference even = _section.XElement.Elements<W.FooterReference>().Where(h => h.Type == W.HeaderFooterValues.Even).FirstOrDefault();
                if (even != null)
                {
                    P.FooterPart footerPart = (P.FooterPart)_doc.Package.MainDocumentPart.GetPartById(even.Id);
                    return new HeaderFooter(_doc, _section, footerPart.Footer, even);
                }
                else
                {
                    if (_doc.Settings.EvenAndOddHeaders
                        && _section.PreviousSection?.HeaderFooters.EvenFooter != null)
                    {
                        return new HeaderFooter(_doc, _section, _section.PreviousSection.HeaderFooters.EvenFooter);
                    }
                    return null;
                }
            }
        }

        #endregion

        #region Public Methods
        /// <summary>
        /// Adds a default header to the section. This is equivalent to <see cref="AddOddHeader()"/>
        /// </summary>
        /// <returns>The added header.</returns>
        public HeaderFooter AddHeader()
        {
            return AddOddHeader();
        }

        /// <summary>
        /// Adds a first page header to the section. If the first page header exists already, no header will be added again. 
        /// And the current first page header will be returned.
        /// </summary>
        /// <returns>The added first page header.</returns>
        public HeaderFooter AddFirstPageHeader()
        {
            if(FirstPageHeader != null)
            {
                return FirstPageHeader;
            }
            string id = RelationshipIdGenerator.Generate(_doc);
            P.HeaderPart hdrPart = PartGenerator.AddNewHeaderPart(_doc, id);
            W.HeaderReference hdrRef = new W.HeaderReference() { Type = W.HeaderFooterValues.First, Id = id };
            _section.XElement.InsertAt(hdrRef, 0);
            return new HeaderFooter(_doc, _section, hdrPart.Header, hdrRef);
        }

        /// <summary>
        /// Adds an odd page header to the section. If the odd page header exists already, no header will be added again. 
        /// And the current odd page header will be returned.
        /// </summary>
        /// <returns>The added odd page header.</returns>
        public HeaderFooter AddOddHeader()
        {
            if (OddHeader != null)
            {
                return OddHeader;
            }
            string id = RelationshipIdGenerator.Generate(_doc);
            P.HeaderPart hdrPart = PartGenerator.AddNewHeaderPart(_doc, id);
            W.HeaderReference hdrRef = new W.HeaderReference() { Type = W.HeaderFooterValues.Default, Id = id };
            _section.XElement.InsertAt(hdrRef, 0);
            return new HeaderFooter(_doc, _section, hdrPart.Header, hdrRef);
        }

        /// <summary>
        /// Adds an even page header to the section. If the even page header exists already, no header will be added again. 
        /// And the current even page header will be returned.
        /// </summary>
        /// <returns>The added even page header.</returns>
        public HeaderFooter AddEvenHeader()
        {
            if (EvenHeader != null)
            {
                return EvenHeader;
            }
            string id = RelationshipIdGenerator.Generate(_doc);
            P.HeaderPart hdrPart = PartGenerator.AddNewHeaderPart(_doc, id);
            W.HeaderReference hdrRef = new W.HeaderReference() { Type = W.HeaderFooterValues.Even, Id = id };
            _section.XElement.InsertAt(hdrRef, 0);
            return new HeaderFooter(_doc, _section, hdrPart.Header, hdrRef);
        }

        /// <summary>
        /// Adds a default footer to the section. This is equivalent to <see cref="AddOddFooter()"/>
        /// </summary>
        /// <returns>The added footer.</returns>
        public HeaderFooter AddFooter()
        {
            return AddOddFooter();
        }

        /// <summary>
        /// Adds a first page footer to the section. If the first page footer exists already, no footer will be added again. 
        /// And the current first page footer will be returned.
        /// </summary>
        /// <returns>The added first page footer.</returns>
        public HeaderFooter AddFirstPageFooter()
        {
            if (FirstPageFooter != null)
            {
                return FirstPageFooter;
            }
            string id = RelationshipIdGenerator.Generate(_doc);
            P.FooterPart ftrPart = PartGenerator.AddNewFooterPart(_doc, id);
            W.FooterReference ftrRef = new W.FooterReference() { Type = W.HeaderFooterValues.First, Id = id };
            _section.XElement.InsertAt(ftrRef, 0);
            return new HeaderFooter(_doc, _section, ftrPart.Footer, ftrRef);
        }

        /// <summary>
        /// Adds an odd page footer to the section. If the odd page footer exists already, no footer will be added again. 
        /// And the current odd page footer will be returned.
        /// </summary>
        /// <returns>The added odd page footer.</returns>
        public HeaderFooter AddOddFooter()
        {
            if (OddFooter != null)
            {
                return OddFooter;
            }
            string id = RelationshipIdGenerator.Generate(_doc);
            P.FooterPart ftrPart = PartGenerator.AddNewFooterPart(_doc, id);
            W.FooterReference ftrRef = new W.FooterReference() { Type = W.HeaderFooterValues.Default, Id = id };
            _section.XElement.InsertAt(ftrRef, 0);
            return new HeaderFooter(_doc, _section, ftrPart.Footer, ftrRef);
        }

        /// <summary>
        /// Adds an even page footer to the section. If the even page footer exists already, no footer will be added again. 
        /// And the current even page footer will be returned.
        /// </summary>
        /// <returns>The added even page footer.</returns>
        public HeaderFooter AddEvenFooter()
        {
            if (EvenFooter != null)
            {
                return EvenFooter;
            }
            string id = RelationshipIdGenerator.Generate(_doc);
            P.FooterPart ftrPart = PartGenerator.AddNewFooterPart(_doc, id);
            W.FooterReference ftrRef = new W.FooterReference() { Type = W.HeaderFooterValues.Even, Id = id };
            _section.XElement.InsertAt(ftrRef, 0);
            return new HeaderFooter(_doc, _section, ftrPart.Footer, ftrRef);
        }

        /// <summary>
        /// Link to the previous section to continue using the same headers and footers if true,
        /// otherwies create new headers or footers for the current section. 
        /// <para>If this is the first section in the document, nothing will happen.</para>
        /// </summary>
        /// <param name="linkToPrevious"></param>
        public void LinkToPrevious(bool linkToPrevious)
        {
            if (_section.PreviousSection == null) return;
            if (FirstPageHeader != null)
                FirstPageHeader.LinkToPrevious = linkToPrevious;
            if (OddHeader != null)
                OddHeader.LinkToPrevious = linkToPrevious;
            if (EvenHeader != null)
                EvenHeader.LinkToPrevious = linkToPrevious;
            if (FirstPageFooter != null)
                FirstPageFooter.LinkToPrevious = linkToPrevious;
            if (OddFooter != null)
                OddFooter.LinkToPrevious = linkToPrevious;
            if (EvenFooter != null)
                EvenFooter.LinkToPrevious = linkToPrevious;
        }

        /// <summary>
        /// Removes all headers and footers that specified in the current section or inherited from the previous section.
        /// </summary>
        public void Remove()
        {
            FirstPageHeader?.Remove();
            OddHeader?.Remove();
            EvenHeader?.Remove();
            FirstPageFooter?.Remove();
            OddFooter?.Remove();
            EvenFooter?.Remove();
        }
        #endregion
    }
}
