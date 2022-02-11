using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using O = DocumentFormat.OpenXml;
using Berry.Docx.Collections;

namespace Berry.Docx
{
    public abstract class DocumentElement : DocumentObject
    {
        public DocumentElement(Document doc, O.OpenXmlElement ele)
            : base(doc, ele)
        {
        }

        public override DocumentObjectCollection ChildObjects
        {
            get;
        }
    }
}
