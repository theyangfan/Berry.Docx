using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Berry.Docx;
using Berry.Docx.Documents;
using Berry.Docx.Field;

using P = DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Test
{
    internal class Test
    {
        public static void Main() {
            string src = @"C:\Users\tomato\Desktop\test.docx";
            string dst = @"C:\Users\tomato\Desktop\dst.docx";

            using (Document doc = new Document(src))
            {
                //Console.WriteLine(doc.Sections[1].HeaderFooters.OddHeader == null);
                foreach (Paragraph p in doc.Sections[1].HeaderFooters.FirstPageHeader.Paragraphs)
                {
                    Console.WriteLine(p.Text);
                }
                //doc.Save();
            }

            //System.Diagnostics.Process.Start(dst);
        }

    }
}
