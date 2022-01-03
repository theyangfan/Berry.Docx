using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;

using OOxml = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using P = DocumentFormat.OpenXml.Packaging;

using Berry.Docx.Documents;
using Berry.Docx.Collections;

namespace Berry.Docx
{
    public class Document
    {
        private P.WordprocessingDocument _doc;
        private Settings _settings;

        /// <summary>
        /// 打开指定文档，如果文档不存在，则创建新文档。
        /// </summary>
        /// <param name="filename">文档名称</param>
        public Document(string filename)
        {
            if (File.Exists(filename))
            {
                _doc = P.WordprocessingDocument.Open(filename, true);
            }
            else
            {
                _doc = DocumentGenerator.Generate(filename);
            }
            _settings = new Settings(_doc.MainDocumentPart.DocumentSettingsPart.Settings);
        }

        /// <summary>
        /// 创建新文档
        /// </summary>
        /// <param name="filename">文档名称</param>
        /// <returns></returns>
        public static Document Create(string filename)
        {
            return new Document(filename);
        }
        /// <summary>
        /// 打开文档
        /// </summary>
        /// <param name="filename">文档名称</param>
        /// <returns></returns>
        public static Document Open(string filename)
        {
            return new Document(filename);
        }

        /// <summary>
        /// 尝试打开文档，不抛出异常
        /// </summary>
        /// <param name="filename">文件名</param>
        /// <param name="isEditable">是否可写</param>
        /// <param name="error">错误信息</param>
        /// <returns></returns>
        public static Document TryOpen(string filename, bool isEditable, out string error)
        {
            error = string.Empty;
            Document doc = null;
            FileInfo file = new FileInfo(filename);
            if (!file.Exists)
            {
                error = $"未找到\"{filename}\"文件，请检查 \"{file.DirectoryName}\" 路径下是否存在该文件!";
                return doc;
            }
            if (file.Length == 0)
            {
                error = $"文件 \"{filename}\" 为空!";
                return doc;
            }
            try
            {
                doc = new Document(filename);
            }
            catch (Exception e)
            {
               error = $"\"{filename}\"读取失败, 请检查是否已关闭该文件!({e.Message})";
            }
            return doc;
        }

        /// <summary>
        /// 保存
        /// </summary>
        public void Save()
        {
            if (_doc != null)
                _doc.Save();
        }
        /// <summary>
        /// 另存为
        /// </summary>
        /// <param name="path"></param>
        public void SaveAs(string path)
        {
            if (_doc != null)
                _doc.SaveAs(path).Close();
        }
        /// <summary>
        /// 关闭文档
        /// </summary>
        public void Close()
        {
            if (_doc != null)
            {
                _doc.Close();
                _doc = null;
            }
        }

        /// <summary>
        /// 文档子类对象
        /// </summary>
        public DocumentObjectCollection ChildObjects
        {
            get
            {
                return new DocumentObjectCollection(ChildObjectsPrivate());
            }
        }

        private IEnumerable<DocumentObject> ChildObjectsPrivate()
        {
            foreach(OOxml.OpenXmlElement ele in _doc.MainDocumentPart.Document.Body.Elements())
            {
                if (ele.GetType() == typeof(W.Paragraph))
                    yield return new Paragraph(ele as W.Paragraph);
                else if (ele.GetType() == typeof(W.Table))
                    yield return new Table(ele as W.Table);
                else
                    yield return new DocumentObject(ele);
            }
        }

        /// <summary>
        /// 返回文档节的集合
        /// </summary>
        public SectionCollection Sections
        {
            get
            {
                return new SectionCollection(SectionsPrivate());
            }
        }

        private IEnumerable<Section> SectionsPrivate()
        {
            foreach (W.SectionProperties sectPr in _doc.MainDocumentPart.Document.Body.Descendants<W.SectionProperties>())
                yield return new Section(sectPr);
        }

        /// <summary>
        /// 最后一节
        /// </summary>
        public Section LastSection
        {
            get
            {
                return new Section(_doc.MainDocumentPart.Document.Body.Elements<W.SectionProperties>().Last());
            }
        }

        /// <summary>
        /// 返回文档的段落集合
        /// </summary>
        public ParagraphCollection Paragraphs
        {
            get
            {
                return new ParagraphCollection(_doc);
            }
        }

        /// <summary>
        /// 返回文档中所有表格
        /// </summary>
        public TableCollection Tables
        {
            get
            {
                return new TableCollection(TablesPrivate());
            }
        }

        private IEnumerable<Table> TablesPrivate()
        {
            foreach (W.Table table in _doc.MainDocumentPart.Document.Body.Elements<W.Table>())
                yield return new Table(table);
        }

        /// <summary>
        /// 返回文档页眉集合
        /// </summary>
        public HeaderCollection Headers
        {
            get
            {
                List<Header> headers = new List<Header>();
                foreach (P.HeaderPart p in _doc.MainDocumentPart.HeaderParts)
                {
                    headers.Add(new Header(p.Header));
                }
                    
                return new HeaderCollection(headers);
            }
        }

        /// <summary>
        /// 返回文档页脚集合
        /// </summary>
        public FooterCollection Footers
        {
            get
            {
                List<Footer> footers = new List<Footer>();
                foreach (P.FooterPart p in _doc.MainDocumentPart.FooterParts)
                {
                    footers.Add(new Footer(p.Footer));
                }

                return new FooterCollection(footers);
            }
        }


        /// <summary>
        /// 返回文档的样式集合
        /// </summary>
        public StyleCollection Styles
        {
            get
            {
                return new StyleCollection(StylesPrivate());
            }
        }

        private IEnumerable<Style> StylesPrivate()
        {
            foreach (W.Style style in _doc.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<W.Style>())
            {
                if (style.Type.Value == W.StyleValues.Paragraph)
                    yield return new ParagraphStyle(style);
                else
                    yield return new Style(style);
            }
        }

        /// <summary>
        /// 返回文档的脚注集合
        /// </summary>
        public List<FootEndnote> FootEndnotes
        {
            get
            {
                List<FootEndnote> fenotes = new List<FootEndnote>();
                if(_doc.MainDocumentPart.FootnotesPart!=null)
                {
                    foreach (W.Footnote fn in _doc.MainDocumentPart.FootnotesPart.Footnotes.Elements<W.Footnote>())
                        fenotes.Add(new FootEndnote(fn));
                }
                return fenotes;
            }
        }
        /// <summary>
        /// 返回文档的尾注集合
        /// </summary>
        public List<FootEndnote> Endnotes
        {
            get
            {
                List<FootEndnote> fenotes = new List<FootEndnote>();
                if(_doc.MainDocumentPart.EndnotesPart!=null)
                {
                    foreach (W.Endnote en in _doc.MainDocumentPart.EndnotesPart.Endnotes.Elements<W.Endnote>())
                        fenotes.Add(new FootEndnote(en));
                }
                return fenotes;
            }
        }
        /// <summary>
        /// 全局设置
        /// </summary>
        public Settings Settings { get => _settings; }

        public Paragraph CreateParagraph()
        {
            W.Paragraph paragraph = new W.Paragraph();
            return new Paragraph(paragraph);
        }

        /// <summary>
        /// 更新域代码
        /// </summary>
        public void UpdateFields()
        {
            if(_doc != null)
            {
                P.DocumentSettingsPart settings = _doc.MainDocumentPart.DocumentSettingsPart;
                W.UpdateFieldsOnOpen updateFields = new W.UpdateFieldsOnOpen();
                updateFields.Val = new OOxml.OnOffValue(true);
                settings.Settings.PrependChild(updateFields);
                settings.Settings.Save();
            }
        }
        /// <summary>
        /// 查找文本内容为 text 的所有段落
        /// </summary>
        /// <param name="text">文本内容</param>
        /// <returns></returns>
        public List<Paragraph> Find(string text)
        {
            List<Paragraph> paras = new List<Paragraph>();
            foreach (W.Paragraph p in _doc.MainDocumentPart.Document.Body.Elements<W.Paragraph>())
            {
                if (p.InnerText.Trim() == text)
                    paras.Add(new Paragraph(p));
            }
            return paras;
        }
        /// <summary>
        /// 返回匹配成功的所有段落
        /// </summary>
        /// <param name="pattern"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        public List<Paragraph> Find(string pattern, RegexOptions options)
        {
            List<Paragraph> paras = new List<Paragraph>();
            foreach (W.Paragraph p in _doc.MainDocumentPart.Document.Body.Elements<W.Paragraph>())
            {
                if (Regex.IsMatch(p.InnerText, pattern, options))
                    paras.Add(new Paragraph(p));
            }
            return paras;
        }
        /// <summary>
        /// 获取目录内容段落列表
        /// </summary>
        /// <returns></returns>
        public List<Paragraph> CatalogParagraphs()
        {
            List<Paragraph> paras = new List<Paragraph>();
            bool begin = false;
            foreach (W.Paragraph p in _doc.MainDocumentPart.Document.Body.Elements<W.Paragraph>())
            {
                if (!begin && Regex.IsMatch(p.InnerText, @"^\s*目\s*录\s*$"))
                {
                    begin = true;
                    continue;
                }
                if (begin)
                {
                    if (p.Descendants<W.SectionProperties>().Count() > 0) break;
                    if (!string.IsNullOrWhiteSpace(p.InnerText.Trim()))
                        paras.Add(new Paragraph(p));
                }
            }
            return paras;
        }
    }
}
