using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using P = DocumentFormat.OpenXml.Packaging;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using M = DocumentFormat.OpenXml.Math;
using V = DocumentFormat.OpenXml.Vml;
using Office = DocumentFormat.OpenXml.Vml.Office;

using Berry.Docx.Formatting;
using Berry.Docx.Field;
using Berry.Docx.Collections;
using Berry.Docx.Utils;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// Represent the paragraph.
    /// </summary>
    public class Paragraph : DocumentItem
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Paragraph _paragraph;
        private readonly ParagraphFormat _pFormat;
        private readonly CharacterFormat _cFormat;
        #endregion

        #region Constructors
        /// <summary>
        /// The paragraph constructor.
        /// </summary>
        /// <param name="doc">The owner document.</param>
        public Paragraph(Document doc) : this(doc, new W.Paragraph())
        {
        }

        internal Paragraph(Document doc, W.Paragraph paragraph) : base(doc, paragraph)
        {
            _doc = doc;
            _paragraph = paragraph;
            _pFormat = new ParagraphFormat(_doc, paragraph);
            _cFormat = new CharacterFormat(_doc, paragraph);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The DocumentObject type.
        /// </summary>
        public override DocumentObjectType DocumentObjectType { get => DocumentObjectType.Paragraph; }

        /// <summary>
        /// The child DocumentObjects of this paragraph.
        /// </summary>
        public override DocumentObjectCollection ChildObjects
        {
            get => new ParagraphItemCollection(_paragraph, ChildObjectsPrivate());
        }

        /// <summary>
        /// The paragraph text.
        /// </summary>
        public string Text
        {
            get
            {
                string text = "";
                bool begin = false;
                bool separate = false;
                foreach (O.OpenXmlElement ele in _paragraph.Descendants())
                {
                    if (ele.GetType().FullName.Equals("DocumentFormat.OpenXml.Wordprocessing.Run"))
                    {
                        W.Run run = (W.Run)ele;
                        if (run.Ancestors<W.SimpleField>().Any())
                            continue;
                        if (run.Elements<W.FieldChar>().Any() && run.Elements<W.FieldChar>().First().FieldCharType != null)
                        {
                            string field_type = run.Elements<W.FieldChar>().First().FieldCharType.ToString();
                            if (field_type == "begin")
                            {
                                begin = true;
                                continue;
                            }
                            if (field_type == "separate")
                            {
                                separate = true;
                                continue;
                            }
                            if (field_type == "end")
                            {
                                begin = false;
                                separate = false;
                                continue;
                            }
                        }
                        if (begin && !separate) continue;
                        foreach (O.OpenXmlElement e in run.Elements())
                        {
                            if (e.GetType().FullName.Equals("DocumentFormat.OpenXml.Wordprocessing.Text"))
                                text += (e as W.Text).Text;
                            if (e.GetType().FullName.Equals("DocumentFormat.OpenXml.Wordprocessing.FieldCode"))
                                text += (e as W.FieldCode).Text;
                            if (e.GetType().FullName.Equals("DocumentFormat.OpenXml.Wordprocessing.NoBreakHyphen"))
                                text += "-";
                        }
                    }
                    else if (ele.GetType().FullName.Equals("DocumentFormat.OpenXml.Wordprocessing.SimpleField"))
                    {
                        text += ele.InnerText;
                    }
                }
                return text;
            }
            set
            {
                _paragraph.RemoveAllChildren<W.Run>();
                W.Run run = RunGenerator.Generate(value);
                _paragraph.AddChild(run);
            }
        }

        /// <summary>
        /// The paragraph format.
        /// </summary>
        public ParagraphFormat Format => _pFormat;

        /// <summary>
        /// The common character format of paragraph.
        /// </summary>
        public CharacterFormat CharacterFormat => _cFormat;

        /// <summary>
        /// The paragraph style. 
        /// </summary>
        public ParagraphStyle Style
        {
            get
            {
                if (_paragraph == null || _paragraph.GetStyle(_doc) == null) return null;
                return new ParagraphStyle(_doc, _paragraph.GetStyle(_doc));
            }
        }

        /// <summary>
        /// Gets the owener section of the current paragraph.
        /// </summary>
        public Section Section
        {
            get
            {
                foreach(Section section in _doc.Sections)
                {
                    if (section.Paragraphs.Contains(this))
                        return section;
                }
                return null;
            }
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Insert a section break with the specified type to the current paragraph. 
        /// <para>The current paragraph must have an owner section, otherwise an exception will be thrown.</para>
        /// </summary>
        /// <param name="type">Type of section break.</param>
        /// <exception cref="NullReferenceException"/>
        /// <returns>The section.</returns>
        public Section InsertSectionBreak(SectionBreakType type)
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
                ele.InsertAfterSelf(startMark);
            }
            else
            {
                _paragraph.InsertAt(startMark, 0);
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
                ele.InsertAfterSelf(startMark);
            }
            else
            {
                _paragraph.InsertAt(startMark, 0);
            }
            _paragraph.Append(endMark);
            _paragraph.Append(referenceRun);
        }

        #endregion

        #region Private Methods
        private IEnumerable<DocumentItem> ChildObjectsPrivate()
        {
            foreach (O.OpenXmlElement ele in _paragraph.ChildElements)
            {
                if (ele.GetType() == typeof(W.Run))
                    yield return new TextRange(_doc, ele as W.Run);
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
                if (_pFormat.NumberingFormat == null) return string.Empty;
                string lvlText = _pFormat.NumberingFormat.Format;
                if (_pFormat.NumberingFormat.Style == W.NumberFormatValues.Decimal)
                    lvlText = lvlText.RxReplace(@"%[0-9]", "1");
                else if (_pFormat.NumberingFormat.Style == W.NumberFormatValues.ChineseCounting
                    || _pFormat.NumberingFormat.Style == W.NumberFormatValues.ChineseCountingThousand)
                    lvlText = lvlText.RxReplace(@"%[0-9]", "一");

                return lvlText;
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
