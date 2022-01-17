using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using O = DocumentFormat.OpenXml;

namespace Berry.Docx.Collections
{
    public class DocumentElementCollection : DocumentObjectCollection
    {
        internal DocumentElementCollection(O.OpenXmlElement owner, IEnumerable<DocumentElement> elements)
            : base(owner, elements)
        {
        }

        internal DocumentElementCollection(O.OpenXmlElement owner)
            : base(owner)
        {
        }
    }
}
