using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Berry.Docx.Documents;
using OOxml = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx
{
    public class Section
    {
        private PageSetup _pageSetup = null;
        private OOxml.SectionProperties _sectPr = null;

        public Section(OOxml.SectionProperties sectPr)
        {
            _pageSetup = new PageSetup(sectPr);
            _sectPr = sectPr;
        }
        /// <summary>
        /// 页面设置
        /// </summary>
        public PageSetup PageSetup { get => _pageSetup; }

        /// <summary>
        /// 页码是否设置章标题
        /// </summary>
        public void setPageNumberChapterStyleShow(bool show = false)
        {
            OOxml.PageNumberType numberType = _sectPr.Elements<OOxml.PageNumberType>().FirstOrDefault();
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
