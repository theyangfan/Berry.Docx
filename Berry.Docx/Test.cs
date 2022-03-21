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
            string src = @"C:\Users\zhailiao123\Desktop\test.docx";
            string dst = @"C:\Users\zhailiao123\Desktop\dst.docx";

            using (Document doc = new Document(src))
            {
                Section s1 = doc.Sections[0];
                //Section s2 = doc.Sections[1];
                //Section s3 = doc.Sections[2];
                Regex rx = new Regex("^关键");

                doc.FindAll(rx);

                //doc.SaveAs(dst);
            }
        }

    }
}
