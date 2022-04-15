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
                /*Paragraph p = doc.LastSection.Paragraphs[0];
                foreach(DocumentObject obj in p.ChildObjects)
                {
                    Console.WriteLine(obj.DocumentObjectType);
                }*/
                doc.FootnoteFormat.RestartRule = FootEndnoteNumberRestartRule.EachSection;
                doc.EndnoteFormat.RestartRule = FootEndnoteNumberRestartRule.EachSection;
                //doc.Sections[0].FootnoteFormat.RestartRule = FootEndnoteNumberRestartRule.EachPage;
                //doc.Sections[0].EndnoteFormat.RestartRule = FootEndnoteNumberRestartRule.EachSection;
                Console.WriteLine(doc.FootnoteFormat.RestartRule);
                Console.WriteLine(doc.EndnoteFormat.RestartRule);
                Console.WriteLine(doc.Sections[0].FootnoteFormat.RestartRule);
                Console.WriteLine(doc.Sections[0].EndnoteFormat.RestartRule);
                Console.WriteLine(doc.Sections[1].FootnoteFormat.RestartRule);
                Console.WriteLine(doc.Sections[1].EndnoteFormat.RestartRule);

                doc.SaveAs(dst);
            }
        }

    }
}
