using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx;
using Berry.Docx.Collections;

namespace Berry.Docx.Documents
{
    public class HeaderFooter
    {
        #region Private Members
        private Document _doc;
        private W.Header _header;
        private W.Footer _footer;
        #endregion

        #region Constructors
        internal HeaderFooter(Document doc, W.Header header)
        {
            _doc = doc;
            _header = header;
        }
        internal HeaderFooter(Document doc, W.Footer footer)
        {
            _doc = doc;
            _footer = footer;
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
