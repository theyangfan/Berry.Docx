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
            string src = @"C:\Users\Tomato\Desktop\test.docx";
            string dst = @"C:\Users\Tomato\Desktop\dst.docx";

            using (Document doc = new Document(src))
            {
                Paragraph p = doc.LastSection.Paragraphs[1];
                PageSetup page = doc.Sections[0].PageSetup;
                //page.CharSpace = 29.3f;
                //page.LineSpace = 10f;
                p.AppendComment("test", "tets");
                doc.SaveAs(dst);
            }
        }
    }
}
