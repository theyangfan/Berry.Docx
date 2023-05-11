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
        public static void Main()
        {
            string src = @"C:\Users\zhailiao123\Desktop\docs\debug\test.docx";
            string dst = @"C:\Users\zhailiao123\Desktop\docs\debug\dst.docx";
            using (Document doc = new Document(src, FileShare.ReadWrite))
            {
                var p = doc.Paragraphs[0];
                var toc = p.AppendTOC(1, 3);
                toc.Update();

                doc.SaveAs(dst);
            }
        }
    }
}
