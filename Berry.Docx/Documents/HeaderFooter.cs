using System;
using System.Collections.Generic;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx;
using Berry.Docx.Collections;

namespace Berry.Docx.Documents
{
    public class HeaderFooter
    {
        #region Private Members
        private Document _doc;
        private Section _section;
        private W.Header _header;
        private W.Footer _footer;
        private string _relationshipID;
        #endregion

        #region Constructors
        internal HeaderFooter(Document doc, Section section, W.Header header, string relationshipID)
        {
            _doc = doc;
            _section = section;
            _header = header;
            _relationshipID = relationshipID;
        }
        internal HeaderFooter(Document doc, Section section, W.Footer footer, string relationshipID)
        {
            _doc = doc;
            _section = section;
            _footer = footer;
            _relationshipID = relationshipID;
        }
        internal HeaderFooter(Document doc, Section section, HeaderFooter headerFooter)
        {
            _doc = doc;
            _section = section;
            _header = headerFooter.Header;
            _footer = headerFooter.Footer;
            _relationshipID = headerFooter.RelationshipID;
        }
        #endregion

        #region Public Properties

        public ParagraphCollection Paragraphs
        {
            get
            {
                if (_header != null)
                    return new ParagraphCollection(_header, GetParagraphs());
                else
                    return new ParagraphCollection(_footer, GetParagraphs());
            }
        }

        public bool LinkToPrevious
        {
            get
            {
                if (_section.PreviousSection == null)
                    return false;
                if (_header != null)
                    return !_section.XElement.Elements<W.HeaderReference>().Where(h => h.Id == _relationshipID).Any();
                else
                    return !_section.XElement.Elements<W.FooterReference>().Where(h => h.Id == _relationshipID).Any();
            }
            set
            {
                if (_section.PreviousSection == null) return;
                if (value)
                {
                    if(_header != null)
                    {
                        W.HeaderReference headerReference = _section.XElement.Elements<W.HeaderReference>()
                            .Where(h => h.Id == _relationshipID).FirstOrDefault();
                        if(headerReference != null)
                        {
                            headerReference.Remove();
                            _doc.Package.MainDocumentPart.DeletePart(_relationshipID);
                        }
                    }
                    if (_footer != null)
                    {
                        W.FooterReference footerReference = _section.XElement.Elements<W.FooterReference>()
                            .Where(h => h.Id == _relationshipID).FirstOrDefault();
                        if (footerReference != null)
                        {
                            footerReference.Remove();
                            _doc.Package.MainDocumentPart.DeletePart(_relationshipID);
                        }
                    }
                }
                else
                {
                }
            }
        }
        #endregion

        #region Public Methods

        #endregion
        internal W.Header Header => _header;
        internal W.Footer Footer => _footer;
        internal string RelationshipID => _relationshipID;

        #region Private Methods
        private IEnumerable<Paragraph> GetParagraphs()
        {
            if(_header != null)
            {
                foreach (var item in _header.Elements<W.Paragraph>())
                {
                    yield return new Paragraph(_doc, item);
                }
            }
            else if(_footer != null)
            {
                foreach (var item in _footer.Elements<W.Paragraph>())
                {
                    yield return new Paragraph(_doc, item);
                }
            }
            
        }
        #endregion
    }
}
