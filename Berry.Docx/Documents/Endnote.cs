using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Collections;

namespace Berry.Docx.Documents
{
    public class Endnote
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Endnote _endnote;
        #endregion

        #region Constructors
        public Endnote(Document doc, W.Endnote endnote)
        {
            _doc = doc;
            _endnote = endnote;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the id of the current endnote.
        /// </summary>
        public int Id => _endnote.Id != null ? (int)_endnote.Id : -1;
        /// <summary>
        /// Gets the paragraph which references the current endnote.
        /// </summary>
        public Paragraph ReferencedParagraph
        {
            get
            {
                foreach (Section section in _doc.Sections)
                {
                    foreach (Paragraph paragraph in section.Paragraphs)
                    {
                        if (paragraph.XElement.Descendants<W.EndnoteReference>().Where(e => e.Id != null && e.Id == Id).Any())
                            return paragraph;
                    }
                }
                return null;
            }
        }
        /// <summary>
        /// Gets endnote paragraphs.
        /// </summary>
        public ParagraphCollection Paragraphs => new ParagraphCollection(_endnote, GetParagraphs());

        #endregion

        #region Private Methods
        private IEnumerable<Paragraph> GetParagraphs()
        {
            foreach(W.Paragraph p in _endnote.Elements<W.Paragraph>())
            {
                yield return new Paragraph(_doc, p);
            }
        }
        #endregion
    }
}
