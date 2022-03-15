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
    public class HeaderFooters
    {
        #region Private Members
        private Document _doc;
        private Section _section;
        #endregion

        #region Constructors
        internal HeaderFooters(Document doc, Section section)
        {
            _doc = doc;
            _section = section;
        }
        #endregion

        #region Public Properties
        public bool DifferentEvenAndOddHeaders
        {
            get => _doc.Settings.EvenAndOddHeaders;
            set => _doc.Settings.EvenAndOddHeaders = value;
        }

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

        public HeaderFooter Header => OddHeader;

        public HeaderFooter FirstPageHeader
        {
            get
            {
                W.HeaderReference first = _section.XElement.Elements<W.HeaderReference>().Where(h => h.Type == W.HeaderFooterValues.First).FirstOrDefault();
                if(first != null)
                {
                    P.HeaderPart headerPart = (P.HeaderPart)_doc.Package.MainDocumentPart.GetPartById(first.Id);
                    return new HeaderFooter(_doc, headerPart.Header);
                }
                else
                {
                    if (_section.XElement.Elements<W.TitlePage>().Any())
                    {
                        return _section.PreviousSection?.HeaderFooters.FirstPageHeader;
                    }
                    return null;
                }
            }
        }

        public HeaderFooter OddHeader
        {
            get
            {
                W.HeaderReference odd = _section.XElement.Elements<W.HeaderReference>().Where(h => h.Type == W.HeaderFooterValues.Default).FirstOrDefault();
                if (odd != null)
                {
                    P.HeaderPart headerPart = (P.HeaderPart)_doc.Package.MainDocumentPart.GetPartById(odd.Id);
                    return new HeaderFooter(_doc, headerPart.Header);
                }
                else
                {
                    return _section.PreviousSection?.HeaderFooters.OddHeader;
                }
            }
        }

        public HeaderFooter EvenHeader
        {
            get
            {
                W.HeaderReference even = _section.XElement.Elements<W.HeaderReference>().Where(h => h.Type == W.HeaderFooterValues.Even).FirstOrDefault();
                if (even != null)
                {
                    P.HeaderPart headerPart = (P.HeaderPart)_doc.Package.MainDocumentPart.GetPartById(even.Id);
                    return new HeaderFooter(_doc, headerPart.Header);
                }
                else
                {
                    if (_doc.Settings.EvenAndOddHeaders)
                    {
                        return _section.PreviousSection?.HeaderFooters.EvenHeader;
                    }
                    return null;
                }
            }
        }
        #endregion

        #region Public Methods
        public HeaderFooter AddFirstPageHeader()
        {
            if(FirstPageHeader != null)
            {
                throw new InvalidOperationException("The first page header of this section already exists.");
            }
            string id = RelationshipIdGenerator.Generate(_doc);
            P.HeaderPart hdrPart = PartGenerator.AddNewHeaderPart(_doc, id);
            W.HeaderReference headerReference = new W.HeaderReference() { Type = W.HeaderFooterValues.First, Id = id };
            _section.XElement.InsertAt(headerReference, 0);
            return new HeaderFooter(_doc, hdrPart.Header);
        }

        public HeaderFooter AddOddHeader()
        {
            if (OddHeader != null)
            {
                throw new InvalidOperationException("The odd header of this section already exists.");
            }
            string id = RelationshipIdGenerator.Generate(_doc);
            P.HeaderPart hdrPart = PartGenerator.AddNewHeaderPart(_doc, id);
            W.HeaderReference headerReference = new W.HeaderReference() { Type = W.HeaderFooterValues.Default, Id = id };
            _section.XElement.InsertAt(headerReference, 0);
            return new HeaderFooter(_doc, hdrPart.Header);
        }

        public HeaderFooter AddEvenHeader()
        {
            if (EvenHeader != null)
            {
                throw new InvalidOperationException("The even header of this section already exists.");
            }
            string id = RelationshipIdGenerator.Generate(_doc);
            P.HeaderPart hdrPart = PartGenerator.AddNewHeaderPart(_doc, id);
            W.HeaderReference headerReference = new W.HeaderReference() { Type = W.HeaderFooterValues.Even, Id = id };
            _section.XElement.InsertAt(headerReference, 0);
            return new HeaderFooter(_doc, hdrPart.Header);
        }
        #endregion
    }
}
