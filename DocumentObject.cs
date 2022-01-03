using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OO = DocumentFormat.OpenXml;
using OW = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Interface;
using Berry.Docx.Collections;
using Berry.Docx.Documents;
using Berry.Docx.Field;

namespace Berry.Docx
{
    /// <summary>
    /// DocumentObject Class.
    /// </summary>
    public class DocumentObject
    {
        private OO.OpenXmlElement _object = null;

        /// <summary>
        /// DocumentObject
        /// </summary>
        /// <param name="obj"></param>
        public DocumentObject(OO.OpenXmlElement obj)
        {
            _object = obj;
        }

        internal OO.OpenXmlElement OpenXmlElement { get => _object; }
        
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
            get => new DocumentObjectCollection(ChildObjectsPrivate());
        }

        private IEnumerable<DocumentObject> ChildObjectsPrivate()
        {
            foreach (OO.OpenXmlElement ele in _object.ChildElements)
            {
                if (ele.GetType() == typeof(OW.Paragraph))
                    yield return new Paragraph(ele as OW.Paragraph);
                else if (ele.GetType() == typeof(OW.Table))
                    yield return new Table(ele as OW.Table);
                else if (ele.GetType() == typeof(OW.Run))
                    yield return new TextRange(ele as OW.Run);
                else
                    yield return new DocumentObject(ele);
            }
        }

        /// <summary>
        /// 获取之前的同级对象
        /// </summary>
        public DocumentObject PreviousSibling
        {
            get
            {
                if (_object == null || _object.PreviousSibling() == null) return null;
                return new DocumentObject(_object.PreviousSibling());
            }

        }
        /// <summary>
        /// 获取之后的同级对象
        /// </summary>
        public DocumentObject NextSibling
        {
            get
            {
                if (_object == null || _object.NextSibling() == null) return null;
                return new DocumentObject(_object.NextSibling());
            }
        }

        public DocumentObject LastChild
        {
            get
            {
                if (_object == null || _object.LastChild == null) return null;
                return new DocumentObject(_object.LastChild);
            }
        }
    }
}
