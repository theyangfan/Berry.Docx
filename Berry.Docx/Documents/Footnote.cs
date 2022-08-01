using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Collections;
using Berry.Docx.Field;

namespace Berry.Docx.Documents
{
    public class Footnote
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Footnote _footnote;
        #endregion

        #region Constructors
        internal Footnote(Document doc, W.Footnote footnote)
        {
            _doc = doc;
            _footnote = footnote;
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
                if (_footnote.Id != null) return (int)_footnote.Id;
                return -1;
            }
        }

        public FootnoteReference Reference
        {
            get
            {
                foreach (Section section in _doc.Sections)
                {
                    foreach (Paragraph paragraph in section.Paragraphs)
                    {
                        foreach(FootnoteReference fnRef in paragraph.ChildObjects.OfType<FootnoteReference>())
                        {
                            if (fnRef.Id == Id)
                                return fnRef;
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
            foreach (W.Paragraph p in _footnote.Elements<W.Paragraph>())
            {
                yield return new Paragraph(_doc, p);
            }
        }
        #endregion
    }
}
