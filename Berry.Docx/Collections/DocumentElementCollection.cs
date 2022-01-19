using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Collections
{
    public class DocumentElementCollection : DocumentObjectCollection
    {
        private O.OpenXmlElement _owner;
        private IEnumerable<DocumentElement> _elements;
        internal DocumentElementCollection(O.OpenXmlElement owner, IEnumerable<DocumentElement> elements)
            : base(elements)
        {
            _owner = owner;
            _elements = elements;
        }

        public override void Add(DocumentObject obj)
        {
            var newElement = obj.XElement;
            if (_elements.Count() > 0)
            {
                var lastElement = _elements.Last().XElement;
                // 末尾段落包含分节符
                if (lastElement is W.Paragraph && lastElement.Descendants<W.SectionProperties>().Any())
                {
                    // 若包含文本，则在分节符后插入，并将分节符移至插入的段落中
                    if (lastElement.Elements<W.Run>().Any())
                    {
                        if(newElement is W.Paragraph)
                        {
                            var sectPr = lastElement.Descendants<W.SectionProperties>().First();
                            sectPr.Remove();
                            W.Paragraph newParagraph = newElement as W.Paragraph;
                            if (newParagraph.ParagraphProperties == null)
                                newParagraph.ParagraphProperties = new W.ParagraphProperties();
                            newParagraph.ParagraphProperties.AddChild(sectPr);
                            lastElement.InsertAfterSelf(newElement);
                        }
                        else
                        {
                            var LastElementTemp = lastElement.CloneNode(true);
                            LastElementTemp.Descendants<W.SectionProperties>().First().Remove();
                            lastElement.InsertBeforeSelf(LastElementTemp);
                            lastElement.RemoveAllChildren<W.Run>();
                            lastElement.InsertBeforeSelf(newElement);
                        }
                    }
                    else
                    {
                        // 若只包含分节符，则在分节符前插入
                        lastElement.InsertBeforeSelf(newElement);
                    }
                }
                else
                {
                    // 若不包含分节符，则在末尾段落之后插入
                    lastElement.InsertAfterSelf(newElement);
                }
            }
            else
            {
                if (_owner is W.Body)
                {
                    _owner.InsertBefore(newElement, _owner.LastChild);
                    return;
                }
                _owner.AppendChild(newElement);
            }
        }

        public override void InsertAt(DocumentObject obj, int index)
        {

        }
        public override void Remove(DocumentObject obj)
        {

        }
        public override void RemoveAt(int index)
        {

        }
        public override void Clear()
        {

        }
    }
}
