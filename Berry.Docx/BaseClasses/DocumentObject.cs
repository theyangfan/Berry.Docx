using System;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Collections;
using Berry.Docx.Documents;
using Berry.Docx.Field;

namespace Berry.Docx
{
    /// <summary>
    /// Represent a base class that all document objects derive from.
    /// </summary>
    public abstract class DocumentObject : IEquatable<DocumentObject>
    {
        #region Private Members
        private Document _doc;
        private O.OpenXmlElement _object;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the DocumentObject class using the supplied underlying OpenXmlElement.
        /// </summary>
        /// <param name="ownerDoc">Owner document</param>
        /// <param name="ele">Underlying OpenXmlElement</param>
        public DocumentObject(Document ownerDoc, O.OpenXmlElement ele)
        {
            _doc = ownerDoc;
            _object = ele;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the owner document.
        /// </summary>
        public Document Document => _doc;
        /// <summary>
        /// Gets all the child objects of the current object.
        /// </summary>
        public abstract DocumentObjectCollection ChildObjects { get; }
        /// <summary>
        /// Gets the type value of the current object.
        /// </summary>
        public abstract DocumentObjectType DocumentObjectType { get; }

        /// <summary>
        /// Gets the object that immediately precedes the current object. 
        /// </summary>
        public virtual DocumentObject PreviousSibling
        {
            get
            {
                O.OpenXmlElement prev = _object.PreviousSibling();
                if (prev == null) return null;
                if (prev is W.Paragraph)
                    return new Paragraph(_doc, (W.Paragraph)prev);
                if (prev is W.Table)
                    return new Table(_doc, (W.Table)prev);
                if (prev is W.Run)
                    return new TextRange(_doc, (W.Run)prev);
                return null;
            }
        }
        /// <summary>
        /// Gets the object that immediately follows the current object.
        /// </summary>
        public virtual DocumentObject NextSibling
        {
            get
            {
                O.OpenXmlElement next = _object.NextSibling();
                if (next == null) return null;
                if (next is W.Paragraph)
                    return new Paragraph(_doc, (W.Paragraph)next);
                if (next is W.Table)
                    return new Table(_doc, (W.Table)next);
                if (next is W.Run)
                    return new TextRange(_doc, (W.Run)next);
                return null;
            }
        }
        #endregion

        #region Public Operators
        /// <summary>
        /// 
        /// </summary>
        /// <param name="lhs"></param>
        /// <param name="rhs"></param>
        /// <returns></returns>
        public static bool operator ==(DocumentObject lhs, DocumentObject rhs)
        {
            if (ReferenceEquals(lhs, rhs)) return true;
            if (((object)lhs == null) || (object)rhs == null) return false;
            return lhs.XElement == rhs.XElement;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="lhs"></param>
        /// <param name="rhs"></param>
        /// <returns></returns>
        public static bool operator !=(DocumentObject lhs, DocumentObject rhs)
        {
            return !(lhs == rhs);
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        protected abstract DocumentObject Clone();

        /// <summary>
        /// Inserts the specified object immediately before the current object.
        /// </summary>
        /// <param name="obj">The new object to insert.</param>
        public virtual void InserBeforeSelf(DocumentObject obj)
        {
            _object.InsertBeforeSelf(obj.XElement);
        }

        /// <summary>
        /// Inserts the specified object immediately after the current object.
        /// </summary>
        /// <param name="obj">The new object to insert.</param>
        public virtual void InsertAfterSelf(DocumentObject obj)
        {
            _object.InsertAfterSelf(obj.XElement);
        }
        /// <summary>
        /// Indicates whether the current object is equal to another object of the same type.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public bool Equals(DocumentObject obj)
        {
            return this == obj;
        }
        /// <summary>
        /// Indicates whether the current object is equal to another object of the same type.
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            return this == (DocumentObject)obj;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
        #endregion

        #region Internal Properties
        internal O.OpenXmlElement XElement => _object;
        #endregion

        #region Internal Methods
        public virtual void Remove()
        {
            _object.Remove();
        }
        #endregion
    }
}
