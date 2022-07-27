using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Drawing;

using Berry.Docx;
using Berry.Docx.Documents;
using Berry.Docx.Field;
using Berry.Docx.Formatting;

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
                //Paragraph p = doc.LastSection.Paragraphs[0];

                var p = doc.LastSection.AddParagraph();
                p.Text = "121";
                //doc.SaveAs(dst);
            }
        }
    }
}
