using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;
using Berry.Docx.Collections;

namespace Berry.Docx
{
    /// <summary>
    /// Represent the section of document.
    /// </summary>
    public class Section
    {
        #region Private Members
        private Document _document;
        private W.SectionProperties _sectPr;
        private PageSetup _pageSetup;
        #endregion

        #region Constructors
        internal Section(Document document, W.SectionProperties sectPr)
        {
            _document = document;
            _sectPr = sectPr;
            _pageSetup = new PageSetup(document, sectPr);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The Page Layout setup.
        /// </summary>
        public PageSetup PageSetup => _pageSetup;

        /// <summary>
        /// Gets a collection of all child objects in the current section.
        /// </summary>
        public DocumentObjectCollection ChildObjects => new DocumentItemCollection(_document.Package.GetBody(), ChildItems());

        /// <summary>
        /// Gets a collection of all paragraphs in the current section.
        /// </summary>
        public ParagraphCollection Paragraphs => new ParagraphCollection(_document.Package.GetBody(), ChildItems().OfType<Paragraph>());

        /// <summary>
        /// Gets a collection of all tables in the current section.
        /// </summary>
        public TableCollection Tables => new TableCollection(_document.Package.GetBody(), ChildItems().OfType<Table>());
        #endregion

        #region Public Methods

        /// <summary>
        /// Add a new paragraph to the end of section.
        /// </summary>
        /// <returns>The paragraph</returns>
        public Paragraph AddParagraph()
        {
            Paragraph paragraph = new Paragraph(_document);
            ChildObjects.Add(paragraph);
            return paragraph;
        }

        /// <summary>
        /// Add a new Table to the end of section.
        /// </summary>
        /// <param name="rowCnt">Table row count</param>
        /// <param name="columnCnt">Table column count</param>
        /// <returns>The table</returns>
        public Table AddTable(int rowCnt, int columnCnt)
        {
            Table table = new Table(_document, rowCnt, columnCnt);
            ChildObjects.Add(table);
            return table;
        }
        #endregion

        #region Internal Properties
        internal W.SectionProperties XElement => _sectPr;
        #endregion

        #region Private Methods
        /// <summary>
        /// Gets the DocuemntItems between current section and previous section.
        /// </summary>
        /// <returns></returns>
        private IEnumerable<DocumentItem> ChildItems()
        {
            List<O.OpenXmlElement> allElements = _document.Package.GetBody().Elements().ToList();
            int startIndex = 0;
            int endIndex = 0;

            int curentSectIndex = _document.Sections.IndexOf(this);
            // Get index of the first item in the current section 
            if (curentSectIndex > 0)
            {
                Section prevSection = _document.Sections[curentSectIndex - 1];
                startIndex = allElements.FindIndex(
                    e => e.Descendants<W.SectionProperties>().Contains(prevSection.XElement)) + 1;
            }
            // Get index of the last item in the current section
            endIndex = allElements.FindIndex(
                e => e == _sectPr || e.Descendants<W.SectionProperties>().Contains(_sectPr));

            for (int i = startIndex; i <= endIndex; ++i)
            {
                O.OpenXmlElement ele = allElements[i];
                if (ele is W.Paragraph)
                {
                    yield return new Paragraph(_document, (W.Paragraph)ele);
                }
                else if (ele is W.Table)
                {
                    yield return new Table(_document, (W.Table)ele);
                }
            }
        }
        #endregion
    }
}
