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
                ParagraphStyle s = p.GetStyle();
                //TextRange tr = p.ChildItems[2] as TextRange;

                Console.WriteLine(p.Format.GetBeforeSpacing());
                //Console.WriteLine(p.Format.GetRightIndent());
                //Console.WriteLine(p.Format.GetSpecialIndentation());
                Console.WriteLine($"Style: {s.Name}");
                Console.WriteLine(s.ParagraphFormat.GetBeforeSpacing());
                //Console.WriteLine(s.ParagraphFormat.GetRightIndent());
                //Console.WriteLine(s.ParagraphFormat.GetSpecialIndentation());
                Console.WriteLine($"Base Style: {s.BaseStyle.Name}");
                Console.WriteLine(s.BaseStyle.ParagraphFormat.GetBeforeSpacing());
                //Console.WriteLine(s.BaseStyle.ParagraphFormat.GetRightIndent());
                //Console.WriteLine(s.BaseStyle.ParagraphFormat.GetSpecialIndentation());

                doc.SaveAs(dst);
            }
        }
    }
}
