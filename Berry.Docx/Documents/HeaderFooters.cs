using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using P = DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx;
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
                    else
                    {
                        return OddHeader;
                    }
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
                    else
                    {
                        return OddHeader;
                    }
                }
            }
        }
        #endregion

    }
}
