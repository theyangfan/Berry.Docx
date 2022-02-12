using System.Collections.Generic;
using O = DocumentFormat.OpenXml;
using Berry.Docx.Documents;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// Represent a Paragraph collection.
    /// </summary>
    public class ParagraphCollection : DocumentItemCollection
    {
        #region Constructors
        internal ParagraphCollection(O.OpenXmlElement owner, IEnumerable<Paragraph> paragraphs)
            : base(owner, paragraphs)
        {
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the paragraph at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        /// <returns>The paragraph at the specified index.</returns>
        public new Paragraph this[int index] => (Paragraph)base[index];
        #endregion

        #region Public Methods
        /// <summary>
        /// Returns the first paragraph of the current collection.
        /// </summary>
        /// <returns>The first paragraph in the current collection.</returns>
        public new Paragraph First()
        {
            return (Paragraph)base.First();
        }

        /// <summary>
        /// Returns the last paragraph of the current collection.
        /// </summary>
        /// <returns>The last paragraph in the current collection.</returns>
        public new Paragraph Last()
        {
            return (Paragraph)base.Last();
        }

        /// <summary>
        /// Determines whether this collection contains a specified paragraph.
        /// </summary>
        /// <param name="paragraph">The specified paragraph.</param>
        /// <returns>true if the collection contains the specified paragraph; otherwise, false.</returns>
        public bool Contains(Paragraph paragraph)
        {
            return base.Contains(paragraph);
        }

        /// <summary>
        /// Adds the specified paragraph to the end of the current collection.
        /// </summary>
        /// <param name="paragraph">The paragraph instance that was added.</param>
        public void Add(Paragraph paragraph)
        {
            base.Add(paragraph);
        }

        /// <summary>
        /// Searchs for the specified paragraph and returns the zero-based index of the first occurrence within the entire collection.
        /// </summary> 
        /// <param name="paragraph">The specified paragraph.</param>
        /// <returns>The zero-based index of the first occurrence of paragraph within the entire collection,if found; otherwise, -1.</returns>
        public int IndexOf(Paragraph paragraph)
        {
            return base.IndexOf(paragraph);
        }

        /// <summary>
        /// Insert the specified paragraph immediately to the specified index of the current collection.
        /// </summary>
        /// <param name="paragraph">The inserted paragraph instance.</param>
        /// <param name="index">The zero-based index.</param>
        public void InsertAt(Paragraph paragraph, int index)
        {
            base.InsertAt(paragraph, index);
        }

        /// <summary>
        /// Removes the specified paragraph immediately from the current collection.
        /// </summary>
        /// <param name="paragraph"> The paragraph instance that was removed. </param>
        public void Remove(Paragraph paragraph)
        {
            base.Remove(paragraph);
        }
        #endregion

    }
}
