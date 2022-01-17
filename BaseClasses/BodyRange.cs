using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Documents;
using Berry.Docx.Collections;

namespace Berry.Docx
{
    public class BodyRange : DocumentContainer
    {
        private Document _doc;
        private O.OpenXmlElement _owner;
        private DocumentElementCollection _collection;
        internal BodyRange(Document doc, O.OpenXmlElement owner)
            : base(doc, owner)
        {
            _doc = doc;
            _owner = owner;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.BodyRange;

        public override DocumentObjectCollection ChildObjects
        {
            get
            {
                DocumentElementCollection collection = null;
                if (_owner is W.SectionProperties)
                {
                    collection = new DocumentElementCollection(_doc.Package.GetBody(), SectionChildElements());
                }
                return collection;
            }
        }

        private IEnumerable<DocumentElement> SectionChildElements()
        {
            W.SectionProperties sectPr = (W.SectionProperties)_owner;
            List<O.OpenXmlElement> allElements = _doc.Package.GetBody().Elements().ToList();
            List<DocumentElement> elements = new List<DocumentElement>();
            int index = 0;
            if (sectPr == _doc.Package.GetRootSectionProperties())
            {
                index = allElements.Count - 1;
            }
            else
            {
                index = allElements.FindIndex(e => e.Descendants().Contains(sectPr));
            }

            for (int i = index; i >= 0; --i)
            {
                O.OpenXmlElement ele = allElements[i];
                // 保留包含 SectionProperties 元素的段落
                if (i != index && ele.Descendants<W.SectionProperties>().Any())
                    break;
                if(ele is W.Paragraph)
                {
                    elements.Add(new Paragraph(_doc, (W.Paragraph)ele));
                }
                
            }
            elements.Reverse();
            return elements.AsEnumerable();
        }

    }
}
