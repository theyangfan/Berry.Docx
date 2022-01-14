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
    public class Document : IDisposable
    {
        private string _filename = string.Empty;
        private P.WordprocessingDocument _doc;
        private Settings _settings;

        /// <summary>
        /// 打开指定文档，如果文档不存在，则创建新文档。
        /// </summary>
        /// <param name="filename">文档名称</param>
        public Document(string filename)
        {
            _filename = filename;
            if (File.Exists(filename))
            {
                using (P.WordprocessingDocument tempDoc = P.WordprocessingDocument.Open(filename, false))
                {
                    _doc = (P.WordprocessingDocument)tempDoc.Clone();
                }
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
        /// 保存
        /// </summary>
        public void Save()
        {
            SaveAs(_filename);
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
                Dispose();
            }
        }

        public void Dispose()
        {
            _doc.Close();
        }

        public P.WordprocessingDocument Package { get => _doc; }

        /// <summary>
        /// 文档子类对象
        /// </summary>
        public DocumentObjectCollection ChildObjects
        {
            get
            {
                return new DocumentObjectCollection(this, ChildObjectsPrivate());
            }
        }

        private IEnumerable<DocumentObject> ChildObjectsPrivate()
        {
            foreach(OOxml.OpenXmlElement ele in _doc.MainDocumentPart.Document.Body.Elements())
            {
                if (ele.GetType() == typeof(W.Paragraph))
                    yield return new Paragraph(this, ele as W.Paragraph);
                else
                    yield return new DocumentObject(this, ele);
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
                yield return new Section(this, sectPr);
        }

        /// <summary>
        /// 最后一节
        /// </summary>
        public Section LastSection
        {
            get
            {
                return new Section(this, _doc.MainDocumentPart.Document.Body.Elements<W.SectionProperties>().Last());
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
                    yield return new ParagraphStyle(this, style);
                else
                    yield return new Style(this, style);
            }
        }

        /// <summary>
        /// 全局设置
        /// </summary>
        public Settings Settings { get => _settings; }

        public Paragraph CreateParagraph()
        {
            W.Paragraph paragraph = new W.Paragraph();
            return new Paragraph(this, paragraph);
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
                    paras.Add(new Paragraph(this, p));
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
                    paras.Add(new Paragraph(this, p));
            }
            return paras;
        }

        #region Future

        /// <summary>
        /// 更新域代码
        /// </summary>
        private void UpdateFields()
        {
            if (_doc != null)
            {
                P.DocumentSettingsPart settings = _doc.MainDocumentPart.DocumentSettingsPart;
                W.UpdateFieldsOnOpen updateFields = new W.UpdateFieldsOnOpen();
                updateFields.Val = new OOxml.OnOffValue(true);
                settings.Settings.PrependChild(updateFields);
                settings.Settings.Save();
            }
        }

        #endregion

    }
}
