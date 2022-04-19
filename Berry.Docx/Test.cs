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
                Paragraph p = doc.Sections[0].Paragraphs[0];
                foreach(DocumentObject obj in p.ChildObjects)
                {
                    Console.WriteLine(obj.DocumentObjectType);
                }
                p.ChildItems[1].AppendComment("123", "456");
                p.ChildItems[5].AppendComment("456", "789");
                doc.SaveAs(dst);
            }
        }

    }
}
