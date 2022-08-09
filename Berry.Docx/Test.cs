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
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Test
{
    internal class Test
    {
        public static void Main() {
            string src = @"C:\Users\Zhailiao123\Desktop\test\test.docx";
            string dst = @"C:\Users\Zhailiao123\Desktop\test\dst.docx";
            using (Document doc = new Document(src))
            {
                Paragraph p = doc.Paragraphs[0];
                Table table = doc.Tables[1];

                table.Format.FirstRowEnabled = true;
                table.Format.FirstColumnEnabled = true;
                table.Format.LastRowEnabled = true;
                table.Format.LastColumnEnabled = true;


                doc.SaveAs(dst);
            }
        }
    }
}
