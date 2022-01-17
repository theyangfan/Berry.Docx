using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Collections
{
    public class ParagraphItemCollection : DocumentObjectCollection
    {
        private IEnumerable<DocumentObject> _objects;
        internal ParagraphItemCollection(O.OpenXmlElement owner, IEnumerable<DocumentObject> objects):base(owner, objects)
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
    }
}
