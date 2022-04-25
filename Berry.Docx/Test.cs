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
            string src = @"C:\Users\Zhailiao123\Desktop\test\test.docx";
            string dst = @"C:\Users\Zhailiao123\Desktop\test\dst.docx";

            using (Document doc = new Document(src))
            {
                PageSetup page = doc.Sections[0].PageSetup;
                //page.CharSpace = 29.3f;
                //page.LineSpace = 10f;
                page.DocGrid = DocGridType.SnapToChars;
                Console.WriteLine(page.DocGrid);
                Console.WriteLine(page.CharSpace);
                Console.WriteLine(page.LineSpace);
                
                doc.SaveAs(dst);
            }
        }

    }
}
