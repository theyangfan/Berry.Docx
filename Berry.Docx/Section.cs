using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using O = DocumentFormat.OpenXml;
using OW = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;
using Berry.Docx.Collections;

namespace Berry.Docx
{
    public class Section : DocumentContainer
    {
        private Document _document = null;
        
        private OW.SectionProperties _sectPr = null;
        private PageSetup _pageSetup = null;

        private BodyRange _range;

        internal Section(Document document, OW.SectionProperties sectPr)
            : base(document, sectPr)
        {
            _document = document;
            _sectPr = sectPr;
            _pageSetup = new PageSetup(sectPr);
            _range = new BodyRange(document, sectPr);
        }

        public override DocumentObjectCollection ChildObjects
        {
            get => new DocumentElementCollection(_sectPr, null);
        }

        public override DocumentObjectType DocumentObjectType { get => DocumentObjectType.Section; }


        /// <summary>
        /// 页面设置
        /// </summary>
        public PageSetup PageSetup { get => _pageSetup; }

        public BodyRange Range => _range;

        public ParagraphCollection Paragraphs
        {
            get
            {
                return new ParagraphCollection(_document.Package.GetBody(), _range.SectionChildElements<Paragraph>());
            }
        }

        public TableCollection Tables
        {
            get
            {
                return new TableCollection(_document.Package.GetBody(), _range.SectionChildElements<Table>());
            }
        }
        
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

    }
}
