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
                PageSetup page = doc.Sections[0].PageSetup;
                Console.WriteLine(page.PageSize);
                Console.WriteLine(page.Orientation);
                Console.WriteLine(page.Margins);
                Console.WriteLine(page.Gutter);
                Console.WriteLine(page.GutterLocation);
                Console.WriteLine(page.HeaderDistance);
                Console.WriteLine(page.FooterDistance);
                //doc.SaveAs(dst);
            }
        }

    }
}
