using System.Collections.Generic;
using O = DocumentFormat.OpenXml;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// Represent a ParagraphItem collection.
    /// </summary>
    public class ParagraphItemCollection : DocumentItemCollection
    {
        #region Constructors
        internal ParagraphItemCollection(O.OpenXmlElement owner, IEnumerable<DocumentItem> objects) : base(owner, objects)
        {
        }
        #endregion

    }
}
