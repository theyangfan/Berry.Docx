using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Drawing;
using System.Diagnostics;

using Berry.Docx;
using Berry.Docx.Documents;
using Berry.Docx.Field;
using Berry.Docx.Formatting;

using P = DocumentFormat.OpenXml.Packaging;
using O = DocumentFormat.OpenXml;
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
                Paragraph p = doc.Paragraphs[0];
                p.Replace(new Regex(@"(\s*[,，;；､、]\s*|\s+)", RegexOptions.IgnoreCase), ";   ");
                doc.SaveAs(dst);
            }
        }
    }
}
