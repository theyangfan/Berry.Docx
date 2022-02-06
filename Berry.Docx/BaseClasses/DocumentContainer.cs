using Berry.Docx.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using O = DocumentFormat.OpenXml;

namespace Berry.Docx
{
    public abstract class DocumentContainer : DocumentObject
    {
        public DocumentContainer(Document doc, O.OpenXmlElement owner)
            : base(doc, owner)
        {

        }

        public override DocumentObjectCollection ChildObjects { get; }
    }
}
