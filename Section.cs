using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using O = DocumentFormat.OpenXml;
using OW = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;
using Berry.Docx.Collections;

namespace Berry.Docx
{
    public class Section : DocumentContainer
    {
        private Document _document = null;
        
        private OW.SectionProperties _sectPr = null;
        private PageSetup _pageSetup = null;

        private BodyRange _range;

        internal Section(Document document, OW.SectionProperties sectPr)
            : base(document, sectPr)
        {
            _document = document;
            _sectPr = sectPr;
            _pageSetup = new PageSetup(sectPr);
            _range = new BodyRange(document, sectPr);
        }

        public override DocumentObjectCollection ChildObjects
        {
            get => new DocumentElementCollection(_sectPr);
        }

        public override DocumentObjectType DocumentObjectType { get => DocumentObjectType.Section; }


        /// <summary>
        /// 页面设置
        /// </summary>
        public PageSetup PageSetup { get => _pageSetup; }

        public BodyRange Range => _range;

        public ParagraphCollection Paragraphs
        {
            get
            {
                return new ParagraphCollection(_document.Package.GetBody(), ParagraphsPrivate());
            }
        }

        private IEnumerable<Paragraph> ParagraphsPrivate()
        {
            List<OW.Paragraph> all_paragraphs = _document.Package.GetBody().Elements<OW.Paragraph>().ToList();
            List<Paragraph> paragraphs = new List<Paragraph>();
            int index = 0;
            if(_sectPr == _document.Package.GetRootSectionProperties())
            {
                index = all_paragraphs.Count - 1;
            }
            else
            {
                index = all_paragraphs.FindIndex(p => p.Descendants().Contains(_sectPr));
            }

            for (int i = index; i >= 0; --i)
            {
                // 保留包含 SectionProperties 元素的段落
                if (i != index && all_paragraphs[i].Descendants<OW.SectionProperties>().Any())
                    break;
                paragraphs.Add(new Paragraph(_document, all_paragraphs[i]));
            }
            paragraphs.Reverse();
            return paragraphs.AsEnumerable();
        }

    }
}
