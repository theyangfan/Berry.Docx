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

namespace Berry.Docx.Documents
{
    public class Paragraph : DocumentObject
    {
        private Document _doc = null;
        private W.Paragraph _paragraph;
        private ParagraphFormat _pFormat;
        private CharacterFormat _cFormat;

        public Paragraph(Document doc, W.Paragraph paragraph) : base(doc, paragraph)
        {
            _doc = doc;
            _paragraph = paragraph;
            _pFormat = new ParagraphFormat(_doc, paragraph);
            _cFormat = new CharacterFormat(_doc, paragraph);
        }
        
        /// <summary>
        /// 段落文本
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

                W.Run run = new W.Run();
                
                W.RunProperties rPr = new W.RunProperties();
                W.RunFonts rFonts = new W.RunFonts() { Hint = W.FontTypeHintValues.EastAsia };
                rPr.AddChild(rFonts);
                
                W.Text text = new W.Text() { Space = O.SpaceProcessingModeValues.Preserve };
                text.Text = value;

                run.AddChild(rPr);
                run.AddChild(text);

                _paragraph.AddChild(run);
            }
        }

        
        /// <summary>
        /// 段落格式
        /// </summary>
        public ParagraphFormat Format { get => _pFormat; }
        /// <summary>
        /// 段落字符格式
        /// </summary>
        public CharacterFormat CharacterFormat { get => _cFormat; }

        /// <summary>
        /// 样式
        /// </summary>
        public Style Style
        {
            get
            {
                if (_paragraph == null || _paragraph.GetStyle(_doc) == null) return null;
                return new Style(_doc, _paragraph.GetStyle(_doc));
            }
        }

        /// <summary>
        /// 样式名称
        /// </summary>
        public string StyleName
        {
            get
            {
                if (_paragraph == null || _paragraph.GetStyle(_doc) == null) return string.Empty;
                return _paragraph.GetStyle(_doc).StyleName.Val;
            }
        }

        public int CharCount
        {
            get => Text.Length;
        }

        /// <summary>
        /// 从父类集合中移除当前段落
        /// </summary>
        public void Remove()
        {
            if (_paragraph != null) _paragraph.Remove();
        }

        #region Future

        /// <summary>
        /// 段落编号(默认为1)
        /// </summary>
        private string ListText
        {
            get
            {
                if (_pFormat.NumberingFormat == null) return string.Empty;
                string lvlText = _pFormat.NumberingFormat.LevelText;
                if (_pFormat.NumberingFormat.NumberingType == W.NumberFormatValues.Decimal)
                    lvlText = lvlText.RxReplace(@"%[0-9]", "1");
                else if (_pFormat.NumberingFormat.NumberingType == W.NumberFormatValues.ChineseCounting
                    || _pFormat.NumberingFormat.NumberingType == W.NumberFormatValues.ChineseCountingThousand)
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

        /// <summary>
        /// 添加批注
        /// </summary>
        /// <param name="author">作者</param>
        /// <param name="content">内容</param>
        private void AppendComment(string author, string content)
        {
            int id = 0; // 新批注Id
            P.WordprocessingCommentsPart part = _doc.Package.MainDocumentPart.WordprocessingCommentsPart;
            if (part == null)
            {
                part = _doc.Package.MainDocumentPart.AddNewPart<P.WordprocessingCommentsPart>();
                part.Comments = new W.Comments();
            }
            W.Comments comments = part.Comments;
            // Id 值为当前批注最大值加1
            List<int> ids = new List<int>();
            foreach (W.Comment c in comments)
                ids.Add(c.Id.Value.ToInt());
            if (ids.Count > 0)
            {
                ids.Sort();
                id = ids.Last() + 1;
            }
            // 设置批注内容
            W.Paragraph paragraph = new W.Paragraph(new W.Run(new W.Text(content)));
            W.Comment comment = new W.Comment(paragraph) { Id = id.ToString(), Author = author };
            comments.AppendChild(comment);
            // 插入批注标记
            W.CommentRangeStart start = new W.CommentRangeStart() { Id = id.ToString() };
            W.Run run = new W.Run(new W.CommentReference() { Id = id.ToString() });
            W.CommentRangeEnd end = new W.CommentRangeEnd() { Id = id.ToString() };
            _paragraph.InsertAt(start, 0);
            _paragraph.AppendChild(end);
            _paragraph.AppendChild(run);
        }

        #endregion
    }
}
