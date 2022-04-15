// Copyright (c) theyangfan. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;
using Berry.Docx.Collections;
using Berry.Docx.Formatting;

namespace Berry.Docx
{
    /// <summary>
    /// Represent the section of document.
    /// </summary>
    public class Section : IEquatable<Section>
    {
        #region Private Members
        private readonly Document _document;
        private readonly W.SectionProperties _sectPr;
        #endregion

        #region Constructors
        internal Section(Document document, W.SectionProperties sectPr)
        {
            _document = document;
            _sectPr = sectPr;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The Page Layout setup.
        /// </summary>
        public PageSetup PageSetup => new PageSetup(_document, this);

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

        /// <summary>
        /// Gets the previous section.
        /// </summary>
        public Section PreviousSection
        {
            get
            {
                int index = _document.Sections.IndexOf(this);
                if (index > 0)
                    return _document.Sections[index - 1];
                return null;
            }
        }

        /// <summary>
        /// Gets the next section.
        /// </summary>
        public Section NextSection
        {
            get
            {
                int index = _document.Sections.IndexOf(this);
                if (index < _document.Sections.Count - 1)
                    return _document.Sections[index + 1];
                return null;
            }
        }

        /// <summary>
        /// Gets the headers and footers of this section.
        /// </summary>
        public HeaderFooters HeaderFooters => new HeaderFooters(_document, this);

        public FootEndnoteFormat FootnoteFormat => new FootEndnoteFormat(_document, this, NoteType.SectionWideFootnote);

        public FootEndnoteFormat EndnoteFormat => new FootEndnoteFormat(_document, this, NoteType.SectionWideEndnote);

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

        /// <summary>
        /// Indicates whether the current object is equal to another object of the same type.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public bool Equals(Section obj)
        {
            return this == obj;
        }
        /// <summary>
        /// Indicates whether the current object is equal to another object of the same type.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            return this == (Section)obj;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
        #endregion

        #region Public Operators
        /// <summary>
        /// 
        /// </summary>
        /// <param name="lhs"></param>
        /// <param name="rhs"></param>
        /// <returns></returns>
        public static bool operator ==(Section lhs, Section rhs)
        {
            if (ReferenceEquals(lhs, rhs)) return true;
            if (((object)lhs == null) || (object)rhs == null) return false;
            return lhs.XElement == rhs.XElement;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="lhs"></param>
        /// <param name="rhs"></param>
        /// <returns></returns>
        public static bool operator !=(Section lhs, Section rhs)
        {
            return !(lhs == rhs);
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
