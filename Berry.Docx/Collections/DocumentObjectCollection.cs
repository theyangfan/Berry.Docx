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
        private IEnumerable<DocumentObject> _objects;

        /// <summary>
        /// DocumentObject 集合
        /// </summary>
        internal DocumentObjectCollection(IEnumerable<DocumentObject> objects)
        {
            _objects = objects;
        }

        /// <summary>
        /// 返回索引为 index 的 DocumentObject 对象
        /// </summary>
        public DocumentObject this[int index] => _objects.ElementAt(index);

        /// <summary>
        /// 返回集合数量
        /// </summary>
        public virtual int Count => _objects.Count();

        public virtual bool Contains(DocumentObject obj)
        {
            return _objects.Contains(obj);
        }
        public virtual int IndexOf(DocumentObject obj)
        {
            return _objects.ToList().IndexOf(obj);
        }
        public abstract void Add(DocumentObject obj);
        public abstract void InsertAt(DocumentObject obj, int index);
        public abstract void Remove(DocumentObject obj);
        public abstract void RemoveAt(int index);
        public abstract void Clear();

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
