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
            string src = @"C:\Users\tomato\Desktop\test.docx";
            string dst = @"C:\Users\tomato\Desktop\dst.docx";

            using (Document doc = new Document(src))
            {
                //doc.Sections[0].PageSetup.PageSize = new System.Drawing.SizeF(590.9f, 384.1f);
                //doc.Sections[0].PageSetup.Orientation = PageOrientation.Landscape;
                Console.WriteLine(doc.Sections[0].PageSetup.PageSize);
                Console.WriteLine(doc.Sections[0].PageSetup.Orientation);
                Console.WriteLine(doc.Sections[0].PageSetup.Margins);
                Console.WriteLine(doc.Sections[0].PageSetup.Gutter);
                Console.WriteLine(doc.Sections[0].PageSetup.HeaderDistance);
                Console.WriteLine(doc.Sections[0].PageSetup.FooterDistance);
                //doc.SaveAs(dst);
            }
        }

    }
}
