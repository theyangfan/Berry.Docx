// Copyright (c) theyangfan. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using P = DocumentFormat.OpenXml.Packaging;

using Berry.Docx.Documents;
using Berry.Docx.Collections;
using Berry.Docx.Utils;
using Berry.Docx.Field;
using Berry.Docx.Formatting;

namespace Berry.Docx
{
    /// <summary>
    /// Represents a document.
    /// </summary>
    public class Document : IDisposable
    {
        #region Private Members
        private string _filename = string.Empty;
        private Stream _stream = null;
        private MemoryStream _mstream = null;
        private readonly P.WordprocessingDocument _doc;
        private Settings _settings;
        #endregion

        #region Constructor
        /// <summary>
        /// Creates a new instance of the Document class from the specified file. If the file dose not exists, a new file will be created .
        /// </summary>
        /// <param name="filename">Name of the file</param>
        public Document(string filename)
        {
            _filename = filename;
            if (File.Exists(filename))
            {
                // open existing doc
                using (P.WordprocessingDocument tempDoc = P.WordprocessingDocument.Open(filename, false))
                {
                    _doc = (P.WordprocessingDocument)tempDoc.Clone();
                }
            }
            else
            {
                // create new doc
                _doc = DocumentGenerator.Generate(filename);
            }
            _settings = new Settings(this, _doc.MainDocumentPart.DocumentSettingsPart.Settings);
        }
        /// <summary>
        /// Creates a new instance of the Document class from the IO stream.
        /// </summary>
        /// <param name="stream"></param>
        public Document(Stream stream)
        {
            _doc = P.WordprocessingDocument.Open(stream, true);
            _settings = new Settings(this, _doc.MainDocumentPart.DocumentSettingsPart.Settings);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Return a collection of sections in the document.
        /// </summary>
        public SectionCollection Sections => new SectionCollection(SectionsPrivate());

        /// <summary>
        /// Return the last section of the document.
        /// </summary>
        public Section LastSection
        {
            get
            {
                return new Section(this, _doc.MainDocumentPart.Document.Body.Elements<W.SectionProperties>().Last());
            }
        }

        /// <summary>
        /// Return a collection of styles in the document.
        /// </summary>
        public StyleCollection Styles => new StyleCollection(StylesPrivate());

        /// <summary>
        /// Return a collection of footnotes in the document.
        /// </summary>
        public List<FootEndnote> Footnotes
        {
            get
            {
                List<FootEndnote> footnotes = new List<FootEndnote>();
                P.FootnotesPart part = _doc.MainDocumentPart.FootnotesPart;
                if(part != null)
                {
                    foreach(W.Footnote fn in part.Footnotes.Elements<W.Footnote>())
                    {
                        footnotes.Add(new FootEndnote(this, fn));
                    }
                }
                return footnotes;
            }
        }

        /// <summary>
        /// Return a collection of endnotes in the document.
        /// </summary>
        public List<FootEndnote> Endnotes
        {
            get
            {
                List<FootEndnote> endnotes = new List<FootEndnote>();
                P.EndnotesPart part = _doc.MainDocumentPart.EndnotesPart;
                if (part != null)
                {
                    foreach (W.Endnote en in part.Endnotes.Elements<W.Endnote>())
                    {
                        endnotes.Add(new FootEndnote(this, en));
                    }
                }
                return endnotes;
            }
        }

        public FootEndnoteFormat FootnoteFormat => _settings.FootnoteFormt;
        public FootEndnoteFormat EndnoteFormat => _settings.EndnoteFormt;
        #endregion

        #region Internal Settings
        internal Settings Settings { get => _settings; }
        #endregion

        #region Public Methods
        /// <summary>
        /// Create a new paragraph.
        /// </summary>
        /// <returns>The paragraph.</returns>
        public Paragraph CreateParagraph()
        {
            return new Paragraph(this);
        }

        /// <summary>
        /// Create a new table with specified size.
        /// </summary>
        /// <param name="rowCnt">Table row count</param>
        /// <param name="columnCnt">Table Column count</param>
        /// <returns>The table.</returns>
        public Table CreateTable(int rowCnt, int columnCnt)
        {
            return new Table(this, rowCnt, columnCnt);
        }

        /// <summary>
        ///  Searches the document for the first occurrence of the specified regular expression.
        /// </summary>
        /// <param name="pattern">The regular expression to search for a match</param>
        /// <returns>An object that contains information about the match.</returns>
        public TextMatch Find(Regex pattern)
        {
            foreach(Section section in Sections)
            {
                foreach(Paragraph p in section.Paragraphs)
                {
                    Match match = pattern.Match(p.Text);
                    if (match.Success)
                    {
                        return new TextMatch(p, match.Index, match.Index + match.Length - 1);
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Searches the document for all occurrences of a regular expression.
        /// </summary>
        /// <param name="pattern">The regular expression to search for a match</param>
        /// <returns>
        /// A list of the <see cref="TextMatch"/> objects found by the search.
        /// </returns>
        public List<TextMatch> FindAll(Regex pattern)
        {
            List<TextMatch> matches = new List<TextMatch>();
            foreach (Section section in Sections)
            {
                foreach (Paragraph p in section.Paragraphs)
                {
                    foreach(Match match in pattern.Matches(p.Text))
                    {
                        if (match.Success)
                        {
                            matches.Add(new TextMatch(p, match.Index, match.Index + match.Length - 1));
                        }
                    }
                }
            }
            return matches;
        }

        /// <summary>
        /// Save the contents and changes of the docuemnt.
        /// </summary>
        public void Save()
        {
            if (!string.IsNullOrEmpty(_filename))
            {
                SaveAs(_filename);
            }
        }
        /// <summary>
        /// Save the contents and changes to specified file.
        /// </summary>
        /// <param name="filename">Name of file</param>
        public void SaveAs(string filename)
        {
            if (_doc != null && !string.IsNullOrEmpty(filename))
                _doc.SaveAs(filename).Close();
        }

        public void SaveAs(Stream stream)
        {
            if(_doc != null)
            {
                _doc.Save();
                _doc.Clone(stream);
            }
        }

        /// <summary>
        /// Close the document.
        /// </summary>
        public void Close()
        {
            Dispose();
        }

        /// <summary>
        /// Close the document.
        /// </summary>
        public void Dispose()
        {
            _stream?.Close();
            _doc?.Close();
        }
        #endregion

        #region Internal Properties
        internal P.WordprocessingDocument Package => _doc;
#endregion

        #region Private Methods
        private IEnumerable<Section> SectionsPrivate()
        {
            foreach (W.SectionProperties sectPr in _doc.MainDocumentPart.Document.Body.Descendants<W.SectionProperties>())
                yield return new Section(this, sectPr);
        }

        private IEnumerable<Style> StylesPrivate()
        {
            foreach (W.Style style in _doc.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<W.Style>())
            {
                if (style.Type.Value == W.StyleValues.Paragraph)
                    yield return new ParagraphStyle(this, style);
                else
                    yield return new Style(this, style);
            }
        }
#endregion

        #region TODO

        /// <summary>
        /// 更新域代码
        /// </summary>
        private void UpdateFields()
        {
            if (_doc != null)
            {
                P.DocumentSettingsPart settings = _doc.MainDocumentPart.DocumentSettingsPart;
                W.UpdateFieldsOnOpen updateFields = new W.UpdateFieldsOnOpen();
                updateFields.Val = new O.OnOffValue(true);
                settings.Settings.PrependChild(updateFields);
                settings.Settings.Save();
            }
        }
        #endregion

    }
}
