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
        private O.OpenXmlElement _owner;
        private IEnumerable<Paragraph> _paragraphs;
        public ParagraphCollection(O.OpenXmlElement owner, IEnumerable<Paragraph> paragraphs)
        {
            _owner = owner;
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

        public bool Contains(Paragraph paragraph)
        {
            return _paragraphs.Contains(paragraph);
        }

        /// <summary>
        /// 在集合末尾添加段落
        /// </summary>
        /// <param name="paragraph">段落</param>
        public void Add(Paragraph paragraph)
        {
            W.Paragraph newParagraph = paragraph.XElement as W.Paragraph;
            if (_paragraphs.Count() == 0)
            {
                if(_owner is W.Body)
                {
                    _owner.InsertBefore(newParagraph, _owner.LastChild);
                    return;
                }
                _owner.AppendChild(newParagraph);
            }
            else
            {
                W.Paragraph lastParagraph = _paragraphs.Last().XElement as W.Paragraph;
                // 末尾段落包含分节符
                if (lastParagraph.Descendants<W.SectionProperties>().Any())
                {
                    // 若包含文本，则在分节符后插入，并将分节符移至插入的段落中
                    if (lastParagraph.Elements<W.Run>().Any())
                    {
                        W.SectionProperties sectPr = lastParagraph.Descendants<W.SectionProperties>().First();
                        sectPr.Remove();
                        if (newParagraph.ParagraphProperties == null)
                            newParagraph.ParagraphProperties = new W.ParagraphProperties();
                        newParagraph.ParagraphProperties.AddChild(sectPr);
                        lastParagraph.InsertAfterSelf(newParagraph);
                    }
                    else
                    {
                        // 若只包含分节符，则在分节符前插入
                        lastParagraph.InsertBeforeSelf(newParagraph);
                    }
                }
                else
                {
                    // 若不包含分节符，则在末尾段落之后插入
                    lastParagraph.InsertAfterSelf(newParagraph);
                }
            }
        }

        /// <summary>
        /// 返回段落在集合中从零开始的索引
        /// </summary>
        /// <param name="paragraph">段落</param>
        /// <returns></returns>
        public int IndexOf(Paragraph paragraph)
        {
            return _paragraphs.ToList().IndexOf(paragraph);
        }

        /// <summary>
        /// 在集合指定位置插入段落
        /// </summary>
        /// <param name="paragraph">段落</param>
        /// <param name="index">段落位置，从零开始的索引</param>
        public void InsertAt(Paragraph paragraph, int index)
        {
            W.Paragraph newParagraph = paragraph.XElement as W.Paragraph;
            if (_paragraphs.Count() == 0)
            {
                if (index == 0)
                {
                    Add(paragraph);
                }
                else
                {
                    throw new ArgumentOutOfRangeException("index", index, "索引超出范围, 必须为非负值并小于集合大小。");
                }
            }
            else
            {
                if(index == _paragraphs.Count())
                {
                    Add(paragraph);
                }
                else
                {
                    _paragraphs.ElementAt(index).XElement.InsertBeforeSelf(newParagraph);
                }
            }
                
        }

        /// <summary>
        /// 移除段落
        /// </summary>
        /// <param name="paragraph">段落</param>
        public void Remove(Paragraph paragraph)
        {
            if (!Contains(paragraph)) return;
            if (paragraph.XElement.Descendants<W.SectionProperties>().Any())
            {
                paragraph.XElement.RemoveAllChildren<W.Run>();
            }
            else
            {
                paragraph.Remove();
            }
        }

        /// <summary>
        /// 移除指定位置处的段落
        /// </summary>
        /// <param name="index">段落位置，从零开始的索引</param>
        public void RemoveAt(int index)
        {
            Remove(_paragraphs.ElementAt(index));
        }

        /// <summary>
        /// 移除所有段落
        /// </summary>
        public void Clear()
        {
            foreach (Paragraph paragraph in _paragraphs)
                Remove(paragraph);
        }

        public IEnumerator GetEnumerator()
        {
            return _paragraphs.GetEnumerator();
        }

        
    }
}
