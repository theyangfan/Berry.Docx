using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
            string src = @"C:\Users\tomato\Desktop\test.docx";
            string dst = @"C:\Users\tomato\Desktop\dst.docx";

            using (Document doc = new Document(src, FileShare.ReadWrite))
            {
                var paragraph = doc.LastSection.Paragraphs[0];
                var pic = paragraph.ChildItems[0] as Picture;
                Console.WriteLine(pic.Width);
                Console.WriteLine(pic.Height);
                // 保存
                //doc.SaveAs(dst);
            }
        }
    }
}
