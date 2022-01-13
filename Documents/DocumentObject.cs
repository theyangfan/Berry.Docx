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
    public class DocumentObject
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

        public Document Document { get => _doc; }

        internal OO.OpenXmlElement XElement { get => _object; }
        
        /// <summary>
        /// 当前对象的类型
        /// </summary>
        public DocumentObjectType DocumentObjectType
        {
            get
            {
                if (_object == null) return DocumentObjectType.Invalid;
                Type type = _object.GetType();
                if (type == typeof(OW.Paragraph))
                    return DocumentObjectType.Paragraph;
                else if (type == typeof(OW.Table))
                    return DocumentObjectType.Table;
                else if (type == typeof(OW.SectionProperties))
                    return DocumentObjectType.Section;
                else if (type == typeof(OW.Run))
                    return DocumentObjectType.TextRange;
                return DocumentObjectType.Invalid;
            }
        }

        public DocumentObjectCollection ChildObjects
        {
            get => new DocumentObjectCollection(_doc, ChildObjectsPrivate());
        }

        private IEnumerable<DocumentObject> ChildObjectsPrivate()
        {
            foreach (OO.OpenXmlElement ele in _object.ChildElements)
            {
                if (ele.GetType() == typeof(OW.Paragraph))
                    yield return new Paragraph(_doc, ele as OW.Paragraph);
                else if (ele.GetType() == typeof(OW.Run))
                    yield return new TextRange(_doc, ele as OW.Run);
                else
                    yield return new DocumentObject(_doc, ele);
            }
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

        public override bool Equals(object obj)
        {
            return this == (DocumentObject)obj;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

    }
}
