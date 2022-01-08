using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using P = DocumentFormat.OpenXml.Packaging;

using Berry.Docx.Documents;

namespace Berry.Docx.Collections
{
    public class ParagraphCollection : IEnumerable
    {
        private O.OpenXmlElement _container = null;
        private IEnumerable<Paragraph> _paragraphs;
        public ParagraphCollection(O.OpenXmlElement container, IEnumerable<Paragraph> paragraphs)
        {
            _container = container;
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
            if (_paragraphs.Count() == 0)
                _container.AppendChild(paragraph.XElement);
            else
                _paragraphs.Last().XElement.InsertAfterSelf(paragraph.XElement);
        }

        public void Insert(Paragraph paragraph, int index)
        {
            if (_paragraphs.Count() == 0)
                _container.AppendChild(paragraph.XElement);
            else
                _paragraphs.ElementAt(index).XElement.InsertBeforeSelf(paragraph.XElement);
        }

        public IEnumerator GetEnumerator()
        {
            return _paragraphs.GetEnumerator();
        }

        
    }
}
