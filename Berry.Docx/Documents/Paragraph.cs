// Copyright (c) theyangfan. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using P = DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wpc = DocumentFormat.OpenXml.Office2010.Word.DrawingCanvas;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = DocumentFormat.OpenXml.Math;
using V = DocumentFormat.OpenXml.Vml;
using Office = DocumentFormat.OpenXml.Vml.Office;

using Berry.Docx.Formatting;
using Berry.Docx.Field;
using Berry.Docx.Collections;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// Represent the paragraph.
    /// <para>该类表示 Word 文档的一种基本元素：段落。</para>
    /// <para>Paragraph 类支持创建空白段落实例。通过该类可以读写段落的文本内容，段落格式，
    /// 同时支持访问段落中的各类子元素：字符，图片，图形，图表，嵌入式对象等。</para>
    /// <para>通过该类的 AppendSectionBreak 函数可以向段落中插入分节符，以达到文档分节的目的，如下所示：</para>
    /// <example>
    /// <code>
    /// using (Document doc = new Document("example.docx"))
    /// {
    ///      // 在文档末尾插入一个“下一页”分节符
    ///      Paragraph p = doc.LastSection.Paragraphs.Last();
    ///      p.AppendSectionBreak(SectionBreakType.NextPage);
    ///      doc.Save();
    /// }
    /// </code>
    /// </example>
    /// </summary>
    public class Paragraph : DocumentItem
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Paragraph _paragraph;
        private readonly ParagraphFormat _pFormat;
        private readonly CharacterFormat _cFormat;
        private readonly ListFormat _listFormat;
        #endregion

        #region Constructors
        /// <summary>
        /// The paragraph constructor.
        /// </summary>
        /// <param name="doc">The owner document.</param>
        public Paragraph(Document doc) : this(doc, new W.Paragraph())
        {
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="paragraph"></param>
        internal Paragraph(Document doc, W.Paragraph paragraph) : base(doc, paragraph)
        {
            _doc = doc;
            _paragraph = paragraph;
            _pFormat = new ParagraphFormat(doc, paragraph);
            _cFormat = new CharacterFormat(doc, paragraph);
            _listFormat = new ListFormat(doc, paragraph);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The DocumentObject type.
        /// </summary>
        public override DocumentObjectType DocumentObjectType { get => DocumentObjectType.Paragraph; }

        /// <summary>
        ///Gets all child <see cref="ParagraphItem"/> of this paragraph.
        /// </summary>
        public ParagraphItemCollection ChildItems => new ParagraphItemCollection(_paragraph, ChildObjectsPrivate());

        /// <summary>
        /// Gets all child <see cref="DocumentObject"/> of this paragraph. 
        /// </summary>
        public override DocumentObjectCollection ChildObjects => new ParagraphItemCollection(_paragraph, ChildObjectsPrivate());

        /// <summary>
        /// The paragraph text.
        /// </summary>
        public string Text
        {
            get
            {
                StringBuilder text = new StringBuilder();
                foreach(DocumentObject item in ChildObjects)
                {
                    if(item is TextRange)
                    {
                        text.Append(((TextRange)item).Text);
                    }
                }
                return text.ToString();
            }
            set
            {
                _paragraph.RemoveAllChildren<W.Run>();
                W.Run run = RunGenerator.Generate(value);
                _paragraph.AddChild(run);
            }
        }

        /// <summary>
        /// Gets the paragraph format.
        /// </summary>
        public ParagraphFormat Format => _pFormat;

        /// <summary>
        /// Gets the list format.
        /// </summary>
        public ListFormat ListFormat => _listFormat;

        /// <summary>
        /// Gets the character format of paragraph mark for this paragraph.
        /// <para>获取段落标记的字符格式.</para>
        /// </summary>
        public CharacterFormat MarkFormat => _cFormat;

        /// <summary>
        /// Gets the owener section of the current paragraph.
        /// </summary>
        internal Section Section
        {
            get
            {
                foreach(Section section in _doc.Sections)
                {
                    if (section.Paragraphs.Contains(this))
                        return section;
                    foreach(SdtBlock sdt in section.ChildObjects.OfType<SdtBlock>())
                    {
                        if (sdt.SdtContent.ChildObjects.OfType<Paragraph>().Contains(this))
                            return section;
                    }
                }
                return null;
            }
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// The paragraph style. 
        /// </summary>
        public ParagraphStyle GetStyle()
        {
            if (_paragraph == null || _paragraph.GetStyle(_doc) == null) return null;
            return new ParagraphStyle(_doc, _paragraph.GetStyle(_doc));
        }

        /// <summary>
        /// Apply the paragraph style with the specified name to the current paragraph.
        /// <para>为当前段落应用指定名称的段落样式.</para>
        /// </summary>
        /// <param name="styleName">The style name.</param>
        public void ApplyStyle(string styleName)
        {
            if (_paragraph == null || string.IsNullOrEmpty(styleName)) return;
            // 如果为内置样式，则应用内置样式
            if (Style.NameToBuiltIn(styleName) != BuiltInStyle.None)
            {
                ApplyStyle(Style.NameToBuiltIn(styleName));
                return;
            }
            var style = _doc.Styles.FindByName(styleName, StyleType.Paragraph);
            if (style == null)
            {
                style = new ParagraphStyle(_doc, styleName);
                _doc.Styles.Add(style);
            }
            if (_paragraph.ParagraphProperties == null)
                _paragraph.ParagraphProperties = new W.ParagraphProperties();
            _paragraph.ParagraphProperties.ParagraphStyleId = new W.ParagraphStyleId() { Val = style.StyleId };
        }

        /// <summary>
        /// Apply the specified built-in style to the current paragraph.
        /// <para>为当前段落应用指定的内置样式.</para>
        /// </summary>
        /// <param name="bstyle">The built-in style type.</param>
        public void ApplyStyle(BuiltInStyle bstyle)
        {
            if (_paragraph == null) return;
            var style = ParagraphStyle.CreateBuiltInStyle(bstyle, _doc);
            if (style != null)
            {
                if (bstyle == BuiltInStyle.Normal)
                {
                    if (_paragraph.ParagraphProperties?.ParagraphStyleId != null)
                        _paragraph.ParagraphProperties.ParagraphStyleId = null;
                }
                else
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    _paragraph.ParagraphProperties.ParagraphStyleId = new W.ParagraphStyleId() { Val = style.StyleId };
                }
            }
        }
        /// <summary>
        /// Append a section break with the specified type to the current paragraph. 
        /// <para>The current paragraph must have an owner section, otherwise an exception will be thrown.</para>
        /// <para>在当前段落结尾添加一个分节符。当前段落必须在节中，否则会抛出一个异常.</para>
        /// </summary>
        /// <param name="type">Type of section break.</param>
        /// <exception cref="NullReferenceException"/>
        /// <returns>The section.</returns>
        public Section AppendSectionBreak(SectionBreakType type)
        {
            if (Section != null)
            {
                W.SectionProperties curSectPr = Section.XElement;
                // Clone a new SectionProperties from current section.
                W.SectionProperties newSectPr = (W.SectionProperties)curSectPr.CloneNode(true);
                // Set the current section type
                W.SectionType curSectType = curSectPr.Elements<W.SectionType>().FirstOrDefault();
                if (curSectType == null)
                {
                    curSectType = new W.SectionType();
                    curSectPr.AddChild(curSectType);
                }
                switch (type)
                {
                    case SectionBreakType.Continuous:
                        curSectType.Val = W.SectionMarkValues.Continuous;
                        break;
                    case SectionBreakType.OddPage:
                        curSectType.Val = W.SectionMarkValues.OddPage;
                        break;
                    case SectionBreakType.EvenPage:
                        curSectType.Val = W.SectionMarkValues.EvenPage;
                        break;
                    default:
                        curSectType.Remove();
                        break;
                }
                // Move current section to the next new paragraph if SectionProperties is present in current paragraph.
                if (_paragraph.Descendants<W.SectionProperties>().Any())
                {
                    W.Paragraph paragraph = new W.Paragraph() { ParagraphProperties = new W.ParagraphProperties() };
                    curSectPr.Remove();
                    paragraph.ParagraphProperties.AddChild(curSectPr);
                    _paragraph.InsertAfterSelf(paragraph);
                }
                // Add the new SectionProperties to the ParagraphProperties of the current paragraph
                if (_paragraph.ParagraphProperties == null)
                    _paragraph.ParagraphProperties = new W.ParagraphProperties();
                _paragraph.ParagraphProperties.AddChild(newSectPr);

                return new Section(_doc, newSectPr);
            }
            else
            {
                throw new NullReferenceException("The owner section of the current paragraph is null.");
            }
        }

        /// <summary>
        /// Append a Break to the end of the current paragraph.
        /// </summary>
        /// <param name="breakType">The BreakType.</param>
        /// <returns>The Break.</returns>
        public Break AppendBreak(BreakType breakType)
        {
            Break br = new Break(_doc, breakType);
            if (breakType == BreakType.TextWrapping)
                br.Clear = BreakTextRestartLocation.All;
            this.ChildItems.Add(br);
            return br;
        }

        /// <summary>
        ///  Searches the paragraph for the first occurrence of the specified regular expression.
        /// </summary>
        /// <param name="pattern">The regular expression to search for a match</param>
        /// <returns>An object that contains information about the match.</returns>
        public TextMatch Find(Regex pattern)
        {
            Match match = pattern.Match(Text);
            if (match.Success)
            {
                return new TextMatch(this, match.Index, match.Index + match.Length - 1);
            }
            return null;
        }

        /// <summary>
        /// Searches the paragraph for all occurrences of a regular expression.
        /// </summary>
        /// <param name="pattern">The regular expression to search for a match</param>
        /// <returns>
        /// A list of the <see cref="TextMatch"/> objects found by the search.
        /// </returns>
        public List<TextMatch> FindAll(Regex pattern)
        {
            List<TextMatch> matches = new List<TextMatch>();
            foreach (Match match in pattern.Matches(Text))
            {
                if (match.Success)
                {
                    matches.Add(new TextMatch(this, match.Index, match.Index + match.Length - 1));
                }
            }
            return matches;
        }

        /// <summary>
        /// Appends a comment to the current paragraph.
        /// </summary>
        /// <param name="author">The author of the comment.</param>
        /// <param name="contents">The paragraphs content of the comment.</param>
        public void AppendComment(string author, params string[] contents)
        {
            int id = 0; // comment id
            P.WordprocessingCommentsPart part = _doc.Package.MainDocumentPart.WordprocessingCommentsPart;
            if (part == null)
            {
                part = _doc.Package.MainDocumentPart.AddNewPart<P.WordprocessingCommentsPart>();
                part.Comments = new W.Comments();
            }
            W.Comments comments = part.Comments;
            // max id + 1
            List<int> ids = new List<int>();
            foreach (W.Comment c in comments)
                ids.Add(c.Id.Value.ToInt());
            if (ids.Count > 0)
            {
                ids.Sort();
                id = ids.Last() + 1;
            }
            // comments content
            
            W.Comment comment = new W.Comment() { Id = id.ToString(), Author = author };
            foreach(string content in contents)
            {
                W.Paragraph paragraph = new W.Paragraph(new W.Run(new W.Text(content)));
                comment.Append(paragraph);
            }
            comments.Append(comment);
            // comment mark
            W.CommentRangeStart startMark = new W.CommentRangeStart() { Id = id.ToString() };
            W.CommentRangeEnd endMark = new W.CommentRangeEnd() { Id = id.ToString() };
            W.Run referenceRun = new W.Run(new W.CommentReference() { Id = id.ToString() });
            // Insert comment mark
            O.OpenXmlElement ele = _paragraph.FirstChild;
            if(ele is W.ParagraphProperties || ele is W.CommentRangeStart)
            {
                while(ele.NextSibling() != null && ele.NextSibling() is W.CommentRangeStart)
                {
                    ele = ele.NextSibling();
                }
                // Exclude page break
                if(ele.NextSibling() is W.Run 
                    && ele.NextSibling().Elements<W.Break>().Where(b => b.Type == W.BreakValues.Page).Any())
                {
                    ele = ele.NextSibling();
                }
                ele.InsertAfterSelf(startMark);
            }
            else
            {
                int index = 0;
                // Exclude page break
                if (ele is W.Run
                    && ele.Elements<W.Break>().Where(b => b.Type == W.BreakValues.Page).Any())
                {
                    index = 1;
                }
                _paragraph.InsertAt(startMark, index);
            }
            _paragraph.Append(endMark);
            _paragraph.Append(referenceRun);
        }

        /// <summary>
        /// Appends a comment to the current paragraph.
        /// </summary>
        /// <param name="author">The author of the comment.</param>
        /// <param name="contents">The paragraphs content of the comment.</param>
        public void AppendComment(string author, IEnumerable<string> contents)
        {
            int id = 0; // comment id
            P.WordprocessingCommentsPart part = _doc.Package.MainDocumentPart.WordprocessingCommentsPart;
            if (part == null)
            {
                part = _doc.Package.MainDocumentPart.AddNewPart<P.WordprocessingCommentsPart>();
                part.Comments = new W.Comments();
            }
            W.Comments comments = part.Comments;
            // max id + 1
            List<int> ids = new List<int>();
            foreach (W.Comment c in comments)
                ids.Add(c.Id.Value.ToInt());
            if (ids.Count > 0)
            {
                ids.Sort();
                id = ids.Last() + 1;
            }
            // comments content
            W.Comment comment = new W.Comment() { Id = id.ToString(), Author = author };
            foreach (string content in contents)
            {
                W.Paragraph paragraph = new W.Paragraph(new W.Run(new W.Text(content)));
                comment.Append(paragraph);
            }
            comments.Append(comment);
            // comment mark
            W.CommentRangeStart startMark = new W.CommentRangeStart() { Id = id.ToString() };
            W.CommentRangeEnd endMark = new W.CommentRangeEnd() { Id = id.ToString() };
            W.Run referenceRun = new W.Run(new W.CommentReference() { Id = id.ToString() });
            // Insert comment mark
            O.OpenXmlElement ele = _paragraph.FirstChild;
            if (ele is W.ParagraphProperties || ele is W.CommentRangeStart)
            {
                while (ele.NextSibling() != null && ele.NextSibling() is W.CommentRangeStart)
                {
                    ele = ele.NextSibling();
                }
                // Exclude page break
                if (ele.NextSibling() is W.Run
                    && ele.NextSibling().Elements<W.Break>().Where(b => b.Type == W.BreakValues.Page).Any())
                {
                    ele = ele.NextSibling();
                }
                ele.InsertAfterSelf(startMark);
            }
            else
            {
                int index = 0;
                // Exclude page break
                if (ele is W.Run
                    && ele.Elements<W.Break>().Where(b => b.Type == W.BreakValues.Page).Any())
                {
                    index = 1;
                }
                _paragraph.InsertAt(startMark, index);
            }
            _paragraph.Append(endMark);
            _paragraph.Append(referenceRun);
        }

        #endregion

        #region Private Methods
        private IEnumerable<ParagraphItem> ChildObjectsPrivate()
        {
            foreach (O.OpenXmlElement ele in _paragraph.ChildElements)
            {
                if (ele is W.Run)
                {
                    foreach (ParagraphItem item in RunItems((W.Run)ele))
                        yield return item;
                }
                else if(ele is W.Hyperlink)
                {
                    foreach (O.OpenXmlElement e in ele.ChildElements)
                    {
                        if (e is W.Run)
                        {
                            foreach (ParagraphItem item in RunItems((W.Run)e))
                                yield return item;
                        }
                    }
                }
                else if(ele is M.OfficeMath) // Office Math
                {
                    yield return new OfficeMath(_doc, ele as M.OfficeMath);
                }
                else if(ele is M.Paragraph)
                {
                    foreach (M.OfficeMath oMath in ele.Elements<M.OfficeMath>())
                        yield return new OfficeMath(_doc, oMath);
                }
            }
        }

        private IEnumerable<ParagraphItem> RunItems(W.Run run)
        {
            // text range
            if (run.Elements<W.Text>().Any())
                yield return new TextRange(_doc, run);

            // footnote reference
            if (run.Elements<W.FootnoteReference>().Any())
            {
                yield return new FootnoteReference(_doc, run, run.Elements<W.FootnoteReference>().First());
            }
            // endnote reference
            if (run.Elements<W.EndnoteReference>().Any())
            {
                yield return new EndnoteReference(_doc, run, run.Elements<W.EndnoteReference>().First());
            }
            // break
            if (run.Elements<W.Break>().Any())
            {
                yield return new Break(_doc, run, run.Elements<W.Break>().First());
            }
            // drawing
            foreach (W.Drawing drawing in run.Descendants<W.Drawing>())
            {
                A.GraphicData graphicData = drawing.Descendants<A.GraphicData>().FirstOrDefault();
                if (graphicData != null)
                {
                    if (graphicData.FirstChild is Pic.Picture)
                        yield return new Picture(_doc, run, drawing);
                    else if (graphicData.FirstChild is Wps.WordprocessingShape)
                        yield return new Shape(_doc, run, drawing);
                    else if (graphicData.FirstChild is Wpg.WordprocessingGroup)
                        yield return new GroupShape(_doc, run, drawing);
                    else if (graphicData.FirstChild is Wpc.WordprocessingCanvas)
                        yield return new Canvas(_doc, run, drawing);
                    else if (graphicData.FirstChild is Dgm.RelationshipIds)
                        yield return new Diagram(_doc, run, drawing);
                    else if (graphicData.FirstChild is C.ChartReference)
                        yield return new Chart(_doc, run, drawing);
                }
            }
            // vml picture
            if (run.Elements<W.Picture>().Any())
            {
                yield return new Picture(_doc, run, run.Elements<W.Picture>().First());
            }
            // embedded object
            foreach (W.EmbeddedObject obj in run.Elements<W.EmbeddedObject>())
            {
                yield return new EmbeddedObject(_doc, run, obj);
            }
        }
        #endregion

        #region TODO

        /// <summary>
        /// 段落编号(默认为1)
        /// </summary>
        private string ListText
        {
            get
            {
                /*
                if (_pFormat.NumberingFormat == null) return string.Empty;
                string lvlText = _pFormat.NumberingFormat.Format;
                //Console.WriteLine($"{lvlText},{_pFormat.NumberingFormat.Style}");
                if (_pFormat.NumberingFormat.Style == W.NumberFormatValues.Decimal)
                    lvlText = lvlText.RxReplace(@"%[0-9]", "1");
                else if (_pFormat.NumberingFormat.Style == W.NumberFormatValues.ChineseCounting
                    || _pFormat.NumberingFormat.Style == W.NumberFormatValues.ChineseCountingThousand
                    || _pFormat.NumberingFormat.Style == W.NumberFormatValues.JapaneseCounting)
                    lvlText = lvlText.RxReplace(@"%[0-9]", "一");
                
                return lvlText;*/
                return "";
            }
        }

        private FieldCodeCollection FieldCodes
        {
            get
            {
                List<FieldCode> fieldcodes = new List<FieldCode>();
                List<O.OpenXmlElement> childElements = new List<O.OpenXmlElement>();

                int begin_times = 0;
                int end_times = 0;

                foreach (O.OpenXmlElement ele in _paragraph.Descendants())
                {
                    if (ele.GetType().FullName.Equals("DocumentFormat.OpenXml.Wordprocessing.SimpleField"))
                    {
                        fieldcodes.Add(new FieldCode((W.SimpleField)ele));
                    }
                    else if (ele.GetType().FullName.Equals("DocumentFormat.OpenXml.Wordprocessing.Run"))
                    {
                        W.Run run = (W.Run)ele;
                        if (run.Elements<W.FieldChar>().Any() && run.Elements<W.FieldChar>().First().FieldCharType != null)
                        {
                            string field_type = run.Elements<W.FieldChar>().First().FieldCharType.ToString();
                            if (field_type == "begin") begin_times++;
                            else if (field_type == "end") end_times++;
                        }
                        if (begin_times > 0)
                        {
                            childElements.Add(ele);
                            if (end_times == begin_times)
                            {
                                fieldcodes.Add(new FieldCode(childElements));
                                begin_times = 0;
                                end_times = 0;
                                childElements.Clear();
                            }
                        }
                    }
                }
                return new FieldCodeCollection(fieldcodes);
            }
        }
        #endregion
    }
}
