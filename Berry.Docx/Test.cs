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
                Section s1 = doc.Sections[0];
                Section s2 = doc.Sections[1];
                Section s3 = doc.Sections[2];
                
                s1.HeaderFooters.DifferentEvenAndOddHeaders = true;

                s1.HeaderFooters.DifferentFirstPageHeaders = true;
                s1.HeaderFooters.AddFirstPageHeader().Paragraphs[0].Text = "第1节首页页眉";
                s1.HeaderFooters.AddOddHeader().Paragraphs[0].Text = "第1节奇数页眉";
                s1.HeaderFooters.AddEvenHeader().Paragraphs[0].Text = "第1节偶数页眉";

                s2.HeaderFooters.DifferentFirstPageHeaders = false;
                s3.HeaderFooters.DifferentFirstPageHeaders = true;

                //s3.HeaderFooters.LinkToPrevious(false);
                //s3.HeaderFooters.Remove();
                //s3.HeaderFooters.FirstPageHeader.LinkToPrevious = true;
                //s3.HeaderFooters.OddHeader.LinkToPrevious = true;
                //s3.HeaderFooters.EvenHeader.LinkToPrevious = true;

                doc.SaveAs(dst);
            }
        }

    }
}
