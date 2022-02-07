using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OO = DocumentFormat.OpenXml;
using OW = DocumentFormat.OpenXml.Wordprocessing;
using OP = DocumentFormat.OpenXml.Packaging;

using Berry.Docx.Documents;
using Berry.Docx.Collections;
using Berry.Docx.Field;

namespace Berry.Docx
{
    /// <summary>
    /// DocumentObject Class.
    /// </summary>
    public abstract class DocumentObject : IEquatable<DocumentObject>
    {
        private Document _doc = null;
        private OO.OpenXmlElement _object = null;

        /// <summary>
        /// DocumentObject
        /// </summary>
        /// <param name="obj"></param>
        public DocumentObject(Document doc, OO.OpenXmlElement obj)
        {
            _doc = doc;
            _object = obj;
        }

        public Document Document => _doc;

        internal OO.OpenXmlElement XElement => _object;

        public abstract DocumentObjectCollection ChildObjects { get; }

        /// <summary>
        /// 当前对象的类型
        /// </summary>
        public abstract DocumentObjectType DocumentObjectType { get; }

        internal void Remove()
        {
            _object.Remove();
        }

        public static bool operator ==(DocumentObject lhs, DocumentObject rhs)
        {
            // 如果均为null，或实例相同，返回true
            if (ReferenceEquals(lhs, rhs)) return true;
            // 如果只有一项为null，返回false
            if (((object)lhs == null) || (object)rhs == null) return false;
            // 判断两者 OpenXmlElement 是否相等
            return lhs.XElement == rhs.XElement;
        }

        public static bool operator !=(DocumentObject lhs, DocumentObject rhs)
        {
            return !(lhs == rhs);
        }

        public bool Equals(DocumentObject obj)
        {
            return this == obj;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

    }
}
