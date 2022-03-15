using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
                /*
                doc.Sections[0].HeaderFooters.DifferentEvenAndOddHeaders = true;
                doc.Sections[0].HeaderFooters.DifferentFirstPageHeaders = true;
                Console.WriteLine("奇偶页不同：{0}", doc.Sections[0].HeaderFooters.DifferentEvenAndOddHeaders);
                Console.WriteLine("首页不同：{0}", doc.Sections[0].HeaderFooters.DifferentFirstPageHeaders);
                foreach (Paragraph p in doc.Sections[0].HeaderFooters.Header?.Paragraphs)
                {
                    Console.WriteLine(p.Text);
                }
                */
                doc.Sections[0].HeaderFooters.DifferentFirstPageHeaders = true;
                doc.Sections[0].HeaderFooters.DifferentEvenAndOddHeaders = true;
                HeaderFooter h1 = doc.Sections[0].HeaderFooters.AddOddHeader();
                HeaderFooter h2 = doc.Sections[0].HeaderFooters.AddEvenHeader();
                HeaderFooter h3 = doc.Sections[0].HeaderFooters.AddFirstPageHeader();
                h1.Paragraphs[0].Text = "奇数";
                h2.Paragraphs[0].Text = "偶数";
                h3.Paragraphs[0].Text = "首页";
                doc.SaveAs(dst);
            }

            //System.Diagnostics.Process.Start(dst);
        }

    }
}
