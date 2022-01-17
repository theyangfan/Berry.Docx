using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Berry.Docx.Documents;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// DocumentObject 集合
    /// </summary>
    public abstract class DocumentObjectCollection : IEnumerable
    {
        private O.OpenXmlElement _owner;
        private IEnumerable<DocumentObject> _objects;

        /// <summary>
        /// DocumentObject 集合
        /// </summary>
        internal DocumentObjectCollection(O.OpenXmlElement owner, IEnumerable<DocumentObject> objects)
        {
            _owner = owner;
            _objects = objects;
        }

        internal DocumentObjectCollection(O.OpenXmlElement owner)
        {

        }

        /// <summary>
        /// 返回索引为 index 的 DocumentObject 对象
        /// </summary>
        public DocumentObject this[int index]
        {
            get
            {
                return _objects.ElementAt(index);
            }
        }
        /// <summary>
        /// 返回集合数量
        /// </summary>
        public int Count { get => _objects.Count(); }
        /// <summary>
        /// 返回集合枚举器
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            return _objects.GetEnumerator();
        }
    }
}
