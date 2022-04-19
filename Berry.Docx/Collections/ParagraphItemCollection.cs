using System.Collections.Generic;
using O = DocumentFormat.OpenXml;
using Berry.Docx.Field;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// Represent a ParagraphItem collection.
    /// </summary>
    public class ParagraphItemCollection : DocumentItemCollection
    {
        #region Constructors
        internal ParagraphItemCollection(O.OpenXmlElement owner, IEnumerable<ParagraphItem> objects) : base(owner, objects)
        {
        }
        #endregion

        /// <summary>
        /// Gets the paragraph child item at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        /// <returns>The paragraph child item at the specified index.</returns>
        public new ParagraphItem this[int index] => (ParagraphItem)base[index];


    }
}
