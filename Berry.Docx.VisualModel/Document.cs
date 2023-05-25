using System;
using System.Collections.Generic;
using System.Text;

using Berry.Docx;

namespace Berry.Docx.VisualModel
{
    public class Document
    {
        private List<Page> _pages;
        public Document(Berry.Docx.Document doc)
        {
            _pages = new List<Page>();

            int pageIndex = -1;
            int sIndex = 0;
            foreach(var section in doc.Sections)
            {
                if (sIndex == 0 || section.Type != SectionBreakType.Continuous)
                {
                    _pages.Add(new Page(doc, section));
                    pageIndex++;
                }
                float charSpace = section.PageSetup.CharPitch.ToPixel();
                float lineSpace = section.PageSetup.LinePitch.ToPixel();
                var gridType = section.PageSetup.DocGrid;

                foreach(var paragraph in section.Paragraphs)
                {
                    int lineNumber = 0;
                    if(paragraph.Format.PageBreakBefore && _pages[pageIndex].Paragraphs.Count > 0)
                    {
                        _pages.Add(new Page(doc, section));
                        pageIndex++;
                    }
                    while ( !_pages[pageIndex].TryAppend(paragraph, ref lineNumber))
                    {
                        _pages.Add(new Page(doc, section));
                        pageIndex++;
                    }
                }
            }
        }

        public List<Page> Pages => _pages;
    }
}
