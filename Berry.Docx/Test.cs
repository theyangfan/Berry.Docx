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
            string src = @"C:\Users\Zhailiao123\Desktop\test\test.docx";
            string dst = @"C:\Users\Zhailiao123\Desktop\test\dst.docx";

            using (Document doc = new Document(src))
            {
                Paragraph p = doc.LastSection.Paragraphs.Last();
                p.Format.Tabs.Clear();
                p.GetStyle().ParagraphFormat.Tabs.Clear();
                foreach(Tab tab in p.Format.Tabs)
                {
                    Console.WriteLine(tab.Position);
                    Console.WriteLine(tab.Style);
                    Console.WriteLine(tab.Leader);
                }

                doc.SaveAs(dst);
            }
        }
    }
}
