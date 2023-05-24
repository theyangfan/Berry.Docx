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
using Berry.Docx.Field;
using Berry.Docx.Formatting;
using Berry.Docx.Utils;

namespace Berry.Docx
{
    /// <summary>
    /// Represents a Word document.
    /// <para>
    /// 表示一个 Word 文档，读写 Word 文档的入口。
    /// </para>
    /// <para>Document 类支持创建空白 Word 文档对象实例，或者通过打开指定文件或流(仅支持docx文件或流)创建对象实例。</para>
    /// <para>对文档做出修改后，如果想保存修改的内容，你应该显式调用 Save 或 SaveAs 方法。</para>
    /// <para>通过 Document 对象可以访问文档中的节，样式，脚注尾注等内容。</para>
    /// </summary>
    public class Document : IDisposable
    {
        #region Private Members
        private readonly string _filename = string.Empty;
        private readonly Stream _stream;
        private readonly P.WordprocessingDocument _doc;
        private readonly P.OpenSettings _openSettings;
        private readonly Settings _settings;
        private bool _closeStream = true;
        #endregion

        #region Constructor
        /// <summary>
        /// Creates a new empty instance of the Document class.
        /// <para>创建一个新的空白 Word 文档。</para>
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
        /// <param name="filename">The name of the specified file.</param>
        public Document(string filename):this(filename, FileShare.Read)
        {
        }

        /// <summary>
        /// Creates a new instance of the Document class from the specified file. 
        /// If the file dose not exists, a new file will be created .
        /// <para>
        /// 打开指定文件来创建一个新 Document 实例。如果文件不存在，则会创建一个新文件。
        /// </para>
        /// </summary>
        /// <param name="filename">The name of the specified file.</param>
        /// <param name="share">A System.IO.FileShare value specifying the type of access other threads have
        ///     to the file.</param>
        public Document(string filename, FileShare share)
        {
            _filename = filename;
            if (File.Exists(filename))
            {
                // open existing file
                MemoryStream tempStream = new MemoryStream();
                using (FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read, share))
                {
                    stream.CopyTo(tempStream);
                }
                tempStream.Seek(0, SeekOrigin.Begin);
                // handle malformed hyperlink
                _openSettings = new P.OpenSettings();
                _openSettings.AutoSave = false;
                _openSettings.RelationshipErrorHandlerFactory += (pkg) =>
                {
                    return new MalformedURIHandler();
                };
                _doc = P.WordprocessingDocument.Open(tempStream, true, _openSettings);
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
            // handle malformed hyperlink
            _openSettings = new P.OpenSettings();
            _openSettings.AutoSave = false;
            _openSettings.RelationshipErrorHandlerFactory += (pkg) =>
            {
                return new MalformedURIHandler();
            };
            _doc = P.WordprocessingDocument.Open(temp_stream, true, _openSettings);
            _settings = new Settings(this, _doc.MainDocumentPart.DocumentSettingsPart.Settings);
        }
        #endregion

        #region Public Properties
        public DocumentObjectCollection ChildObjects
        {
            get
            {
                return null;
            }
        }


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
        /// Gets a collection of all <see cref="Paragraph"/> in the current document.
        /// <para>返回当前文档中所有段落的集合。</para>
        /// </summary>
        public ParagraphCollection Paragraphs => new ParagraphCollection(Package.GetBody(), GetAllParagraphs());

        /// <summary>
        /// Gets a collection of all <see cref="Table"/> in the current document.
        /// <para>返回当前文档中所有表格的集合。</para>
        /// </summary>
        public TableCollection Tables => new TableCollection(Package.GetBody(), GetAllTables());

        /// <summary>
        /// Return a collection of <see cref="Style"/> that supports traversal in the document. 
        /// <para>返回当前文档中所有样式的可遍历集合。</para>
        /// </summary>
        public StyleCollection Styles => new StyleCollection(this);

        /// <summary>
        /// Return a collection of <see cref="ListStyle"/> that supports traversal in the document. 
        /// <para>返回当前文档中所有列表样式的可遍历集合。</para>
        /// </summary>
        public ListStyleCollection ListStyles => new ListStyleCollection(this);

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
        /// Returns all of the bookmarks in the document.
        /// </summary>
        public BookmarkCollection Bookmarks => new BookmarkCollection(this);

        /// <summary>
        /// Returns the document default paragraph and character formats.
        /// <para>返回文档默认段落和字符格式。</para>
        /// </summary>
        public DocDefaultFormat DefaultFormat => new DocDefaultFormat(this);
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
            {
                _doc.SaveAs(filename).Close();
            }
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
        /// Sets a Boolean value indicating whether close the source stream when close the document.
        /// </summary>
        /// <param name="close"></param>
        public void SetCloseStream(bool close)
        {
            _closeStream = close;
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
            if(_closeStream)
                _stream?.Close();
            _doc?.Close();
        }

        /// <summary>
        /// 更新域代码
        /// </summary>
        public void UpdateFieldsWhenOpen()
        {
            if (_doc != null)
            {
                P.DocumentSettingsPart settings = _doc.MainDocumentPart.DocumentSettingsPart;
                W.UpdateFieldsOnOpen updateFields = new W.UpdateFieldsOnOpen() { Val = true };
                settings.Settings.PrependChild(updateFields);
                settings.Settings.Save();
            }
        }

        /// <summary>
        /// Removes all of the OpenXmlUnknownElement(SmartTag, etc.)
        /// </summary>
        /// <param name="filename">The document file path.</param>
        public static void Normalize(string filename)
        {
            using(var doc = new Document(filename))
            {
                doc.Normalize();
                doc.Save();
            }
        }

        /// <summary>
        /// Removes all of the OpenXmlUnknownElement(SmartTag, etc.)
        /// </summary>
        /// <param name="stream">The document stream.</param>
        public static void Normalize(Stream stream)
        {
            using (var doc = new Document(stream))
            {
                doc.SetCloseStream(false);
                doc.Normalize();
                doc.Save();
            }
        }
        #endregion

        #region Internal Properties
        internal P.WordprocessingDocument Package => _doc;

        /// <summary>
        /// Returns document settings.
        /// <para>返回文档 settings。</para>
        /// </summary>
        internal Settings Settings { get => _settings; }
        #endregion

        #region Private Methods
        private IEnumerable<Section> SectionsPrivate()
        {
            foreach (W.SectionProperties sectPr in _doc.MainDocumentPart.Document.Body.Descendants<W.SectionProperties>())
                yield return new Section(this, sectPr);
        }

        private IEnumerable<Paragraph> GetAllParagraphs()
        {
            foreach(Section section in Sections)
            {
                foreach(Paragraph paragraph in section.Paragraphs)
                {
                    yield return paragraph;
                }
            }
        }

        private IEnumerable<Table> GetAllTables()
        {
            foreach (Section section in Sections)
            {
                foreach (Table table in section.Tables)
                {
                    yield return table;
                }
            }
        }

        private void Normalize()
        {
            List<O.OpenXmlUnknownElement> unknownElements = new List<O.OpenXmlUnknownElement>();
            foreach (var p in Paragraphs)
            {
                foreach (var unknow in p.XElement.Elements<O.OpenXmlUnknownElement>())
                {
                    // SmartTag
                    if (unknow.LocalName == "smartTag")
                    {
                        // run
                        foreach (var r in unknow.Elements().Where(e => e.LocalName == "r"))
                        {
                            unknow.InsertBeforeSelf(r.CloneNode(true));
                        }
                        unknownElements.Add(unknow);
                    }
                }
            }
            foreach (var unknow in unknownElements) unknow.Remove();
            unknownElements.Clear();
        }
        #endregion

    }
}
