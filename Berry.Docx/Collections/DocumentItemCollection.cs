using System.Linq;
using System.Collections.Generic;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// Represent a DocumentItem collection.
    /// </summary>
    public class DocumentItemCollection : DocumentObjectCollection
    {
        #region Private Members
        private O.OpenXmlElement _owner;
        private IEnumerable<DocumentItem> _items;
        #endregion

        #region Constructors
        internal DocumentItemCollection(O.OpenXmlElement owner, IEnumerable<DocumentItem> items)
            : base(items)
        {
            _owner = owner;
            _items = items;
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Returns the first item of the current collection.
        /// </summary>
        /// <returns>The first item in the current collection.</returns>
        public DocumentItem First()
        {
            return _items.First();
        }

        /// <summary>
        /// Returns the last item of the current collection.
        /// </summary>
        /// <returns>The last item in the current collection.</returns>
        public DocumentItem Last()
        {
            return _items.Last();
        }

        /// <summary>
        /// Adds the specified object to the end of the current collection.
        /// </summary>
        /// <param name="obj">The DocumentObject instance that was added.</param>
        public override void Add(DocumentObject obj)
        {
            var newElement = obj.XElement;
            if (_items.Count() > 0)
            {
                var lastElement = _items.Last().XElement;
                // the last item is paragraph and contains section
                if (lastElement is W.Paragraph && lastElement.Descendants<W.SectionProperties>().Any())
                {
                    // if last item contains text, insert the new item after last item and move the section to the new item
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
                        // if only contains section, insert before last item
                        lastElement.InsertBeforeSelf(newElement);
                    }
                }
                else
                {
                    // if doesn't contain section, insert after last item
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

        /// <summary>
        /// Insert the specified object immediately to the specified index of the current collection.
        /// </summary>
        /// <param name="obj">The inserted DocumentObject instance.</param>
        /// <param name="index">The zero-based index.</param>
        public override void InsertAt(DocumentObject obj, int index)
        {
            if (index == _items.Count())
            {
                Add(obj);
            }
            else
            {
                _items.ElementAt(index).XElement.InsertBeforeSelf(obj.XElement);
            }
        }

        /// <summary>
        /// Removes the specified DocumentObject immediately from the current collection.
        /// </summary>
        /// <param name="obj"> The DocumentObject instance that was removed. </param>
        public override void Remove(DocumentObject obj)
        {
            if (!Contains(obj)) return;
            if (obj.XElement.Descendants<W.SectionProperties>().Any())
            {
                obj.XElement.RemoveAllChildren<W.Run>();
            }
            else
            {
                obj.Remove();
            }
        }

        /// <summary>
        /// Removes the DocumentObject at the zero-based index immediately from the current collection.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        public override void RemoveAt(int index)
        {
            Remove(_items.ElementAt(index));
        }

        /// <summary>
        /// Removes all items of the current collection.
        /// </summary>
        public override void Clear()
        {
            foreach(DocumentObject obj in _items)
            {
                Remove(obj);
            }
        }
        #endregion
    }
}
