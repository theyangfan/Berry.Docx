using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// DocumentObject 集合
    /// </summary>
    public class DocumentObjectCollection : IEnumerable
    {
        private Document _doc = null;
        private IEnumerable<DocumentObject> _objects;
        /// <summary>
        /// DocumentObject 集合
        /// </summary>
        public DocumentObjectCollection(Document doc, IEnumerable<DocumentObject> objects)
        {
            _doc = doc;
            _objects = objects;
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

        private IEnumerable<DocumentObject> ChildObjectsPrivate()
        {
            foreach (OO.OpenXmlElement ele in _object.ChildElements)
            {
                if (ele.GetType() == typeof(OW.Paragraph))
                    yield return new Paragraph(_doc, ele as OW.Paragraph);
                else if (ele.GetType() == typeof(OW.Table))
                    yield return new Table(ele as OW.Table);
                else if (ele.GetType() == typeof(OW.Run))
                    yield return new TextRange(ele as OW.Run);
                else
                    yield return new DocumentObject(_doc, ele);
            }
        }
    }
}
