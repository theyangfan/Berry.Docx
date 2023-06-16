using System;
using System.Collections.Generic;
using System.Text;
using BD = Berry.Docx.Documents;

namespace Berry.Docx.Visual
{
    public class Document
    {
        #region Private Members
        private List<Page> _pages;
        #endregion

        #region Constructors
        public Document(Berry.Docx.Document doc)
        {
            _pages = new List<Page>();

            int pageIndex = -1;
            int sIndex = 0;
            foreach(var section in doc.Sections)
            {
                if (sIndex++ == 0 || section.Type != SectionBreakType.Continuous)
                {
                    _pages.Add(new Page(doc, section));
                    pageIndex++;
                }
                float charSpace = section.PageSetup.CharPitch.ToPixel();
                float lineSpace = section.PageSetup.LinePitch.ToPixel();
                var gridType = section.PageSetup.DocGrid;

                foreach(var obj in section.ChildObjects)
                {
                    if(obj is BD.Paragraph)
                    {
                        var paragraph = (BD.Paragraph)obj;
                        int lineNumber = 0;
                        if (paragraph.Format.PageBreakBefore && _pages[pageIndex].ChildItems.Count > 0)
                        {
                            _pages.Add(new Page(doc, section));
                            pageIndex++;
                        }
                        while (!_pages[pageIndex].TryAppend(paragraph, ref lineNumber))
                        {
                            _pages.Add(new Page(doc, section));
                            pageIndex++;
                        }
                    }
                    else if(obj is BD.Table)
                    {
                        var table = (BD.Table)obj;
                        while (!_pages[pageIndex].TryAppend(table))
                        {
                            _pages.Add(new Page(doc, section));
                            pageIndex++;
                        }
                    }
                }
            }
        }
        #endregion

        #region Public Properties
        public List<Page> Pages => _pages;
        #endregion
    }
}
