using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OW = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;
using Berry.Docx.Collections;

namespace Berry.Docx
{
    public class Section
    {
        private Document _document = null;
        
        private OW.SectionProperties _sectPr = null;
        private PageSetup _pageSetup = null;

        public Section(Document document, OW.SectionProperties sectPr)
        {
            _document = document;
            _pageSetup = new PageSetup(sectPr);
            _sectPr = sectPr;
        }

        /// <summary>
        /// 页面设置
        /// </summary>
        public PageSetup PageSetup { get => _pageSetup; }
        public ParagraphCollection Paragraphs
        {
            get
            {
                return new ParagraphCollection(_document, ParagraphsPrivate());
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
                if (i == index && all_paragraphs[i].Descendants().Contains(_sectPr))
                {
                    if(all_paragraphs[i].Elements<OW.Run>().Any())
                        paragraphs.Add(new Paragraph(_document, all_paragraphs[i]));
                    continue;
                }
                if (all_paragraphs[i].Descendants<OW.SectionProperties>().Any())
                    break;
                paragraphs.Add(new Paragraph(_document, all_paragraphs[i]));
            }
            paragraphs.Reverse();
            return paragraphs.AsEnumerable();
        }

        /// <summary>
        /// 页码是否设置章标题
        /// </summary>
        public void setPageNumberChapterStyleShow(bool show = false)
        {
            OW.PageNumberType numberType = _sectPr.Elements<OW.PageNumberType>().FirstOrDefault();
            if (numberType != null)
            {
                if (numberType.ChapterStyle != null && !show)
                {
                    numberType.ChapterStyle = null;
                }
            }
        }
    }
}
