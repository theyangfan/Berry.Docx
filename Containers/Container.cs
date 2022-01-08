using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using O = DocumentFormat.OpenXml;

namespace Berry.Docx
{
    public class Container : DocumentObject
    {
        public Container(Document document, O.OpenXmlElement xml)
            :base(document, xml)
        {

        }
    }
}
