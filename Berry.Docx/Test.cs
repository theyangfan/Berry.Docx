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
            string src = @"C:\Users\Zhailiao123\Desktop\test\test.docx";
            string dst = @"C:\Users\Zhailiao123\Desktop\test\dst.docx";
            string file = @"C:\Users\Zhailiao123\Desktop\test\test1.jpg";

            using (Document doc = new Document(src, FileShare.ReadWrite))
            {
                Paragraph p = doc.Paragraphs[0];
                p.TextReadingMode = TextReadingMode.IncludeAll;
                Console.WriteLine(p.Text);
                //doc.SaveAs(dst);
            }
        }
    }
}
