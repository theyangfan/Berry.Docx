using System;
using System.Collections.Generic;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx;
using Berry.Docx.Collections;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// /// Represent the header or footer of the section.
    /// </summary>
    public class HeaderFooter
    {
        #region Private Members
        private readonly Document _doc;
        private readonly Section _section;
        private readonly W.Header _header;
        private readonly W.HeaderReference _hdrRef;
        private readonly W.Footer _footer;
        private readonly W.FooterReference _ftrRef;
        #endregion

        #region Constructors
        internal HeaderFooter(Document doc, Section section, W.Header header, W.HeaderReference hdrRef)
        {
            _doc = doc;
            _section = section;
            _header = header;
            _hdrRef = hdrRef;
        }
        internal HeaderFooter(Document doc, Section section, W.Footer footer, W.FooterReference ftrRef)
        {
            _doc = doc;
            _section = section;
            _footer = footer;
            _ftrRef = ftrRef;
        }
        internal HeaderFooter(Document doc, Section section, HeaderFooter headerFooter)
        {
            _doc = doc;
            _section = section;
            _header = headerFooter.Header;
            _hdrRef = headerFooter.HeaderReference;
            _footer = headerFooter.Footer;
            _ftrRef = headerFooter.FooterReference;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the paragraphs of the header or footer.
        /// </summary>
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

        /// <summary>
        /// Link to the previous section to continue using the same header or footer if true,
        /// otherwies create new header or footer for the current section. 
        /// <para>If this is the first section in the document, nothing will happen.</para>
        /// </summary>
        public bool LinkToPrevious
        {
            get
            {
                if (_section.PreviousSection == null)
                    return false;
                if (_header != null)
                    return !_section.XElement.ChildElements.Contains(_hdrRef);
                else
                    return !_section.XElement.ChildElements.Contains(_ftrRef);
            }
            set
            {
                if (_section.PreviousSection == null) return;
                if (value)
                {
                    if(_header != null)
                    {
                        if(_section.XElement.ChildElements.Contains(_hdrRef))
                        {
                            _hdrRef.Remove();
                            _doc.Package.MainDocumentPart.DeletePart(_header.HeaderPart);
                        }
                    }
                    if (_footer != null)
                    {
                        if (_section.XElement.ChildElements.Contains(_ftrRef))
                        {
                            _ftrRef.Remove();
                            _doc.Package.MainDocumentPart.DeletePart(_footer.FooterPart);
                        }
                    }
                }
                else
                {
                    if(_header != null)
                    {
                        if (!_section.XElement.ChildElements.Contains(_hdrRef))
                        {
                            string id = IDGenerator.GenerateRelationshipID(_doc);
                            PartGenerator.AddNewHeaderPart(_doc, id);
                            W.HeaderReference hdrRef = (W.HeaderReference)_hdrRef.CloneNode(true);
                            hdrRef.Id = id;
                            _section.XElement.InsertAt(hdrRef, 0);
                        }
                    }
                    if (_footer != null)
                    {
                        if (!_section.XElement.ChildElements.Contains(_ftrRef))
                        {
                            string id = IDGenerator.GenerateRelationshipID(_doc);
                            PartGenerator.AddNewFooterPart(_doc, id);
                            W.FooterReference ftrRef = (W.FooterReference)_ftrRef.CloneNode(true);
                            ftrRef.Id = id;
                            _section.XElement.InsertAt(ftrRef, 0);
                        }
                    }
                }
            }
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Removes the current header or footer.
        /// </summary>
        public void Remove()
        {
            if(_header != null)
            {
                _hdrRef.Remove();
                _doc.Package.MainDocumentPart.DeletePart(_header.HeaderPart);
            }
            if (_footer != null)
            {
                _ftrRef.Remove();
                _doc.Package.MainDocumentPart.DeletePart(_footer.FooterPart);
            }
        }
        #endregion

        #region Internal Properties
        internal W.Header Header => _header;
        internal W.HeaderReference HeaderReference => _hdrRef;
        internal W.Footer Footer => _footer;
        internal W.FooterReference FooterReference => _ftrRef;
        #endregion

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
