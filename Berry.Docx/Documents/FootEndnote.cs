using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Collections;

namespace Berry.Docx.Documents
{
    public class FootEndnote
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Footnote _footnote;
        private readonly W.Endnote _endnote;
        #endregion

        #region Constructors
        internal FootEndnote(Document doc, W.Footnote footnote)
        {
            _doc = doc;
            _footnote = footnote;
        }
        internal FootEndnote(Document doc, W.Endnote endnote)
        {
            _doc = doc;
            _endnote = endnote;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the id of the current footnote.
        /// </summary>
        public int Id
        {
            get
            {
                if (_footnote?.Id != null) return (int)_footnote.Id;
                if (_endnote?.Id != null) return (int)_endnote.Id;
                return -1;
            }
        }
        /// <summary>
        /// Gets the paragraph which references the current footnote.
        /// </summary>
        public Paragraph ReferencedParagraph
        {
            get
            {
                foreach(Section section in _doc.Sections)
                {
                    foreach(Paragraph paragraph in section.Paragraphs)
                    {
                        if(_footnote != null)
                        {
                            if (paragraph.XElement.Descendants<W.FootnoteReference>().Where(n => n.Id != null && n.Id == Id).Any())
                                return paragraph;
                        }
                        else if(_endnote != null)
                        {
                            if (paragraph.XElement.Descendants<W.EndnoteReference>().Where(n => n.Id != null && n.Id == Id).Any())
                                return paragraph;
                        }
                    }
                }
                return null;
            }
        }
        /// <summary>
        /// Gets footnote paragraphs.
        /// </summary>
        public ParagraphCollection Paragraphs => new ParagraphCollection(_footnote, GetParagraphs());
        #endregion

        #region Private Methods
        private IEnumerable<Paragraph> GetParagraphs()
        {
            if(_footnote != null)
            {
                foreach (W.Paragraph p in _footnote.Elements<W.Paragraph>())
                {
                    yield return new Paragraph(_doc, p);
                }
            }
            else if(_endnote != null)
            {
                foreach (W.Paragraph p in _endnote.Elements<W.Paragraph>())
                {
                    yield return new Paragraph(_doc, p);
                }
            }
        }
        #endregion
    }
}
