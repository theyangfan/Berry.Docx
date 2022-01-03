using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OOxml = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using P = DocumentFormat.OpenXml.Packaging;

using Berry.Docx.Documents;

namespace Berry.Docx.Collections
{
    public class ParagraphCollection : IEnumerable
    {
        private P.WordprocessingDocument _doc = null;
        private IEnumerable<Paragraph> _paragraphs;
        public ParagraphCollection(P.WordprocessingDocument doc)
        {
            _doc = doc;
            _paragraphs = ParagraphsPrivate();
        }

        public Paragraph this[int index]
        {
            get
            {
                return _paragraphs.ElementAt(index);
            }
        }

        /// <summary>
        /// 返回集合数量
        /// </summary>
        public int Count { get => _paragraphs.Count(); }

        public void Add(Paragraph paragraph)
        {
            _doc.MainDocumentPart.Document.Body.AddChild(paragraph.OpenXmlElement);
        }

        public IEnumerator GetEnumerator()
        {
            return _paragraphs.GetEnumerator();
        }

        private IEnumerable<Paragraph> ParagraphsPrivate()
        {
            foreach (W.Paragraph p in _doc.MainDocumentPart.Document.Body.Elements<W.Paragraph>())
                yield return new Paragraph(p);
        }
    }
}
