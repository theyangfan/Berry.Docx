using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Collections;
using Berry.Docx.Field;

namespace Berry.Docx.Documents
{
    public class Endnote
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Endnote _endnote;
        #endregion

        #region Constructors
        internal Endnote(Document doc, W.Endnote endnote)
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
                if (_endnote.Id != null) return (int)_endnote.Id;
                return -1;
            }
        }

        public EndnoteReference Reference
        {
            get
            {
                foreach (Section section in _doc.Sections)
                {
                    foreach (Paragraph paragraph in section.Paragraphs)
                    {
                        foreach (EndnoteReference enRef in paragraph.ChildObjects.OfType<EndnoteReference>())
                        {
                            if (enRef.Id == Id)
                                return enRef;
                        }
                    }
                }
                return null;
            }
        }
        /// <summary>
        /// Gets footnote paragraphs.
        /// </summary>
        public ParagraphCollection Paragraphs => new ParagraphCollection(_endnote, GetParagraphs());
        #endregion

        #region Private Methods
        private IEnumerable<Paragraph> GetParagraphs()
        {
            foreach (W.Paragraph p in _endnote.Elements<W.Paragraph>())
            {
                yield return new Paragraph(_doc, p);
            }
        }
        #endregion
    }
}
