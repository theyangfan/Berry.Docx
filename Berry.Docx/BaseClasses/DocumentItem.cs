using O = DocumentFormat.OpenXml;
using Berry.Docx.Collections;

namespace Berry.Docx
{
    /// <summary>
    /// Represent a base class that all document item objects derive from.
    /// </summary>
    public abstract class DocumentItem : DocumentObject
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the DocumentItem class using the supplied underlying OpenXmlElement.
        /// </summary>
        /// <param name="ownerDoc">Owner document</param>
        /// <param name="ele">Underlying OpenXmlElement</param>
        public DocumentItem(Document ownerDoc, O.OpenXmlElement ele)
            : base(ownerDoc, ele)
        {
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets all the child objects of the current item.
        /// </summary>
        public override DocumentObjectCollection ChildObjects
        {
            get;
        }
        #endregion
    }
}
