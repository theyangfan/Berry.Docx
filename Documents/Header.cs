using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Collections;

namespace Berry.Docx.Documents
{
    public class Header
    {
        private W.Header _header; 

        public Header(W.Header header) 
        {
            _header = header;
        }

        /*
        public ParagraphCollection Paragraphs
        {
            get
            {
                List<Paragraph> paras = new List<Paragraph>();
                foreach (W.Paragraph p in _header.Elements<W.Paragraph>())
                    paras.Add(new Paragraph(p));
                return new ParagraphCollection(paras);
            }
        }
        */

    }


}
