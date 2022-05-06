﻿// Copyright (c) theyangfan. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

/*
 * Document.cs 文件定义了 Document 类，该类是读写 Word 文档的入口。
 * Document 类支持创建空白 Word 文档对象实例，或者通过打开指定文件或流(仅支持docx文件或流)创建对象实例。
 * 对文档做出修改后，如果想保存修改的内容，你应该显式调用 Save 或 SaveAs 方法。
 * 该类实现了 IDisposable 接口，所以你可以在 using 语句中声明对象，而非显式调用 Close 方法，如下所示：
 * // 示例开始
 * using (Document doc = new Document("example.docx"))
 * {
 *      // 一些操作
 *      ...
 *      doc.Save();
 * }
 * // 示例结束
 * 通过 Document 对象可以访问文档中的节，样式，脚注尾注等内容。 
 */

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
using Berry.Docx.Field;
using Berry.Docx.Formatting;

namespace Berry.Docx
{
    /// <summary>
    /// Represents a Word document.
    /// <para>
    /// 表示一个 Word 文档，读写 Word 文档的入口。
    /// </para>
    /// </summary>
    public class Document : IDisposable
    {
        #region Private Members
        private readonly string _filename = string.Empty;
        private readonly Stream _stream;
        private readonly P.WordprocessingDocument _doc;
        private readonly Settings _settings;
        #endregion

        #region Constructor
        /// <summary>
        /// Creates a new empty instance of the Document class.
        /// <para>创建一个新的空白 Document 实例。</para>
        /// </summary>
        public Document()
        {
            _stream = new MemoryStream();
            MemoryStream temp_stream = new MemoryStream();
            _doc = DocumentGenerator.Generate(temp_stream);
            _settings = new Settings(this, _doc.MainDocumentPart.DocumentSettingsPart.Settings);
        }

        /// <summary>
        /// Creates a new instance of the Document class from the specified file. 
        /// If the file dose not exists, a new file will be created .
        /// <para>
        /// 打开指定文件来创建一个新 Document 实例。如果文件不存在，则会创建一个新文件。
        /// </para>
        /// </summary>
        /// <param name="filename">Name of the file</param>
        public Document(string filename)
        {
            _filename = filename;
            if (File.Exists(filename))
            {
                // open existing file
                using (P.WordprocessingDocument tempDoc = P.WordprocessingDocument.Open(filename, false))
                {
                    _doc = (P.WordprocessingDocument)tempDoc.Clone();
                }
            }
            else
            {
                // create new file
                _doc = DocumentGenerator.Generate(filename);
            }
            _settings = new Settings(this, _doc.MainDocumentPart.DocumentSettingsPart.Settings);
        }
        /// <summary>
        /// Creates a new instance of the Document class from the io stream.
        /// The stream must be a valid read-write Word file stream.
        /// <para>
        /// 打开指定流来创建一个新 Document 实例。流必须是一个有效的可读写 Word 文档流。
        /// </para>
        /// </summary>
        /// <param name="stream">The read-write Word stream.</param>
        public Document(Stream stream)
        {
            _stream = stream;
            MemoryStream temp_stream = new MemoryStream();
            stream.CopyTo(temp_stream);
            _doc = P.WordprocessingDocument.Open(temp_stream, true);
            _settings = new Settings(this, _doc.MainDocumentPart.DocumentSettingsPart.Settings);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Return a collection of <see cref="Section"/> that supports traversal in the document. 
        /// <para>返回当前文档中所有节的可遍历集合</para>
        /// </summary>
        public SectionCollection Sections => new SectionCollection(SectionsPrivate());

        /// <summary>
        /// Return the last section of the document.
        /// <para>返回文档的最后一节。</para>
        /// </summary>
        public Section LastSection
        {
            get
            {
                return new Section(this, _doc.MainDocumentPart.Document.Body.Elements<W.SectionProperties>().Last());
            }
        }

        /// <summary>
        /// Return a collection of <see cref="Style"/> that supports traversal in the document. 
        /// <para>返回当前文档中所有样式的可遍历集合。</para>
        /// </summary>
        public StyleCollection Styles => new StyleCollection(this);

        /// <summary>
        /// Return a collection of <see cref="Footnote"/> in the document.
        /// <para>返回当前文档中的所有脚注。</para>
        /// </summary>
        public List<Footnote> Footnotes
        {
            get
            {
                List<Footnote> footnotes = new List<Footnote>();
                P.FootnotesPart part = _doc.MainDocumentPart.FootnotesPart;
                if(part != null)
                {
                    foreach(W.Footnote fn in part.Footnotes.Elements<W.Footnote>())
                    {
                        footnotes.Add(new Footnote(this, fn));
                    }
                }
                return footnotes;
            }
        }

        /// <summary>
        /// Return a collection of <see cref="Endnote"/> in the document.
        /// <para>返回当前文档中的所有尾注。</para>
        /// </summary>
        public List<Endnote> Endnotes
        {
            get
            {
                List<Endnote> endnotes = new List<Endnote>();
                P.EndnotesPart part = _doc.MainDocumentPart.EndnotesPart;
                if (part != null)
                {
                    foreach (W.Endnote en in part.Endnotes.Elements<W.Endnote>())
                    {
                        endnotes.Add(new Endnote(this, en));
                    }
                }
                return endnotes;
            }
        }
        /// <summary>
        /// Returns the footnote format in the document.
        /// <para>返回当前文档的脚注格式。</para>
        /// </summary>
        public FootEndnoteFormat FootnoteFormat => _settings.FootnoteFormt;
        /// <summary>
        /// Returns the endnote format in the document.
        /// <para>返回当前文档的尾注格式。</para>
        /// </summary>
        public FootEndnoteFormat EndnoteFormat => _settings.EndnoteFormt;
        /// <summary>
        /// Returns the document default paragraph and character formats.
        /// <para>返回文档默认段落和字符格式。</para>
        /// </summary>
        public DocDefaultFormat DefaultFormat => new DocDefaultFormat(this);
        #endregion

        #region Internal Settings
        /// <summary>
        /// Returns document settings.
        /// <para>返回文档 settings。</para>
        /// </summary>
        internal Settings Settings { get => _settings; }
        #endregion

        #region Public Methods
        /// <summary>
        /// Create a new paragraph.
        /// <para>新建一个段落。</para>
        /// </summary>
        /// <returns>The paragraph.</returns>
        public Paragraph CreateParagraph()
        {
            return new Paragraph(this);
        }

        /// <summary>
        /// Create a new table with specified size.
        /// <para>新建一个指定尺寸的表格。</para>
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
        ///  <para>查找文档中第一个匹配指定正则表达式的文本。</para>
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
        /// <para>查找文档中所有匹配指定正则表达式的文本。</para>
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
        /// <para>保存文档内容。</para>
        /// </summary>
        public void Save()
        {
            if (!string.IsNullOrEmpty(_filename))
            {
                SaveAs(_filename);
            }
            else if(_stream != null)
            {
                SaveAs(_stream);
            }
        }
        /// <summary>
        /// Save the contents and changes to specified file.
        /// <para>保存文档内容至指定文件。</para>
        /// </summary>
        /// <param name="filename">Name of file</param>
        public void SaveAs(string filename)
        {
            if (_doc != null && !string.IsNullOrEmpty(filename))
                _doc.SaveAs(filename).Close();
        }
        /// <summary>
        /// Save the contents and changes to specified stream.
        /// <para>保存文档内容至指定流中。</para>
        /// </summary>
        /// <param name="stream">The destination stream.</param>
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
        /// <para>关闭文档。</para>
        /// </summary>
        public void Close()
        {
            Dispose();
        }

        /// <summary>
        /// Close the document.
        /// <para>关闭文档。</para>
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
