using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Collections;

namespace Berry.Docx.Documents
{
    public class Footer
    {
        private W.Footer _footer;

        public Footer(W.Footer footer)
        {
            _footer = footer;
        }

        /*
        public ParagraphCollection Paragraphs
        {
            get 
            {
                List<Paragraph> paras = new List<Paragraph>();
                foreach (W.Paragraph p in _footer.Elements<W.Paragraph>())
                    paras.Add(new Paragraph(p));
                return new ParagraphCollection(paras);
            }
        }
        */
    }
}
