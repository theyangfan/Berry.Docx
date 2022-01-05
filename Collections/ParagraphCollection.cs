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
        public ParagraphCollection(IEnumerable<Paragraph> paragraphs)
        {
            _paragraphs = paragraphs;
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
           //_doc.MainDocumentPart.Document.Body.AddChild(paragraph.OpenXmlElement);
        }

        public IEnumerator GetEnumerator()
        {
            return _paragraphs.GetEnumerator();
        }

        
    }
}
