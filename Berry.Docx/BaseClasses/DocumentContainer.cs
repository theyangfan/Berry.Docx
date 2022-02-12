using O = DocumentFormat.OpenXml;
using Berry.Docx.Collections;

namespace Berry.Docx
{
    /// <summary>
    /// Represent a base class that all document container objects derive from.
    /// </summary>
    public abstract class DocumentContainer : DocumentObject
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the DocumentContainer class using the supplied underlying OpenXmlElement.
        /// </summary>
        /// <param name="ownerDoc">Owner document</param>
        /// <param name="ownerEle">Underlying OpenXmlElement</param>
        public DocumentContainer(Document ownerDoc, O.OpenXmlElement ownerEle)
            : base(ownerDoc, ownerEle)
        {

        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets all the child objects of the current item.
        /// </summary>
        public override DocumentObjectCollection ChildObjects { get; }
        #endregion


    }
}
