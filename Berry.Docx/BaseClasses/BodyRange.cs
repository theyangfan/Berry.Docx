using System.Collections.Generic;
using System.Linq;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Documents;
using Berry.Docx.Collections;

namespace Berry.Docx
{
    /// <summary>
    /// Represent a body content range in the document.
    /// </summary>
    public class BodyRange : DocumentContainer
    {
        #region Private Members
        private Document _doc;
        private O.OpenXmlElement _owner;
        #endregion

        #region Constructors
        internal BodyRange(Document doc, O.OpenXmlElement owner)
            : base(doc, owner)
        {
            _doc = doc;
            _owner = owner;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// 
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.BodyRange;
        /// <summary>
        /// 
        /// </summary>
        public override DocumentObjectCollection ChildObjects
        {
            get
            {
                DocumentItemCollection collection = null;
                if (_owner is W.SectionProperties)
                {
                    collection = new DocumentItemCollection(_doc.Package.GetBody(), SectionChildElements());
                }
                return collection;
            }
        }
        #endregion

        #region Internal Methods
        internal IEnumerable<T> SectionChildElements<T>() where T : DocumentItem
        {
            return SectionChildElements().OfType<T>();
        }

        /// <summary>
        /// Gets the DocuemntItems between current section and previous section.
        /// </summary>
        /// <returns></returns>
        internal IEnumerable<DocumentItem> SectionChildElements()
        {
            List<O.OpenXmlElement> allElements = _doc.Package.GetBody().Elements().ToList();
            int startIndex = 0;
            int endIndex = 0;

            Section curSection = new Section(_doc, (W.SectionProperties)_owner);
            int curentSectIndex = _doc.Sections.IndexOf(curSection);
            
            if(curentSectIndex > 0)
            {
                Section prevSection = _doc.Sections[curentSectIndex - 1];
                startIndex = allElements.FindIndex(e => e.Descendants<W.SectionProperties>().Contains(prevSection.XElement));
            }

            endIndex = allElements.FindIndex(e => e == curSection.XElement || e.Descendants<W.SectionProperties>().Contains(curSection.XElement));

            for (int i = startIndex; i <= endIndex; ++i)
            {
                O.OpenXmlElement ele = allElements[i];
                if(ele is W.Paragraph)
                {
                    yield return new Paragraph(_doc, (W.Paragraph)ele);
                }
                else if(ele is W.Table)
                {
                    yield return new Table(_doc, (W.Table)ele);
                }
            }
        }
        #endregion
    }
}
