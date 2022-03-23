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
            string src = @"C:\Users\zhailiao123\Desktop\dst.docx";
            string dst = @"C:\Users\zhailiao123\Desktop\dst.docx";

            using (Document doc = new Document(src))
            {
                /*TextMatch m = doc.Find(new Regex("测"));
                if(m != null)
                {
                    Console.WriteLine(m.GetAsOneRange().CharacterFormat.Bold);
                }*/
                Paragraph p = doc.Sections[0].Paragraphs[2];
                Console.WriteLine(p.CharacterFormat.Bold);
                Console.WriteLine(p.ChildObjects.OfType<TextRange>().Last().CharacterFormat.FontNameEastAsia);
                //doc.SaveAs(dst);
            }
        }

    }
}
