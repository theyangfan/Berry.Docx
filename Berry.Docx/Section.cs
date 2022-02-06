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
    public class Section : DocumentContainer
    {
        #region Private Members
        private Document _document;
        private W.SectionProperties _sectPr;
        private PageSetup _pageSetup;
        private BodyRange _range;
        #endregion

        #region Constructors
        /// <summary>
        /// Create a Section class instance.
        /// </summary>
        /// <param name="document">Owner document object</param>
        public Section(Document document)
            : this(document, document.LastSection.XElement.CloneNode(true) as W.SectionProperties)
        {
        }

        internal Section(Document document, W.SectionProperties sectPr)
            : base(document, sectPr)
        {
            _document = document;
            _sectPr = sectPr;
            _pageSetup = new PageSetup(document, sectPr);
            _range = new BodyRange(document, sectPr);
        }
        #endregion

        #region Public Properties

        /// <summary>
        /// The DocumentObject type.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Section;

        /// <summary>
        /// The Page Layout setup.
        /// </summary>
        public PageSetup PageSetup => _pageSetup;

        /// <summary>
        /// The range of section content.
        /// </summary>
        public BodyRange Range => _range;

        /// <summary>
        /// The paragraphs in this section.
        /// </summary>
        public ParagraphCollection Paragraphs
        {
            get
            {
                return new ParagraphCollection(_document.Package.GetBody(), _range.SectionChildElements<Paragraph>());
            }
        }
        /// <summary>
        /// The tables in this section.
        /// </summary>
        public TableCollection Tables
        {
            get
            {
                return new TableCollection(_document.Package.GetBody(), _range.SectionChildElements<Table>());
            }
        }
        #endregion

        #region Public Methods

        /// <summary>
        /// Add a new paragraph to the end of section.
        /// </summary>
        /// <returns>The paragraph</returns>
        public Paragraph AddParagraph()
        {
            Paragraph paragraph = new Paragraph(_document);
            Paragraphs.Add(paragraph);
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
            Tables.Add(table);
            return table;
        }
        #endregion
    }
}
