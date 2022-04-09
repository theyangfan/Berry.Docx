using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
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
            string src = @"C:\Users\zhailiao123\Desktop\test\test.docx";
            string dst = @"C:\Users\zhailiao123\Desktop\test\dst.docx";

            using (Document doc = new Document(src))
            {
                /*Paragraph p = doc.LastSection.Paragraphs[0];
                foreach(DocumentObject obj in p.ChildObjects)
                {
                    Console.WriteLine(obj.DocumentObjectType);
                }*/
                foreach(Footnote fn in doc.Footnotes)
                {
                    Console.WriteLine(fn.ReferencedParagraph?.Text);
                }
                foreach (Endnote en in doc.Endnotes)
                {
                    Console.WriteLine(en.ReferencedParagraph?.Text);
                }
            }
        }

    }
}
