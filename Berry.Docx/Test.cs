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
            string src = @"C:\Users\Tomato\Desktop\test.docx";
            string dst = @"C:\Users\Tomato\Desktop\dst.docx";

            using (Document doc = new Document(src))
            {
                Paragraph p = doc.LastSection.Paragraphs.Last();
                ParagraphStyle s = p.GetStyle();
                //TextRange tr = p.ChildItems[2] as TextRange;
                p.Format.SetBeforeSpacing(13, SpacingUnit.Point);
                p.Format.SetAfterSpacing(2.5f, SpacingUnit.Line);
                p.Format.SetLineSpacing(2.8F, LineSpacingRule.Multiple);

                Console.WriteLine(p.Format.GetBeforeSpacing());
                Console.WriteLine(p.Format.GetAfterSpacing());
                Console.WriteLine(p.Format.GetLineSpacing());
                Console.WriteLine($"Style: {s.Name}");
                Console.WriteLine(s.ParagraphFormat.GetBeforeSpacing());
                Console.WriteLine(s.ParagraphFormat.GetAfterSpacing());
                Console.WriteLine(s.ParagraphFormat.GetLineSpacing());
                Console.WriteLine($"Base Style: {s.BaseStyle.Name}");
                Console.WriteLine(s.BaseStyle.ParagraphFormat.GetBeforeSpacing());
                Console.WriteLine(s.BaseStyle.ParagraphFormat.GetAfterSpacing());
                Console.WriteLine(s.BaseStyle.ParagraphFormat.GetLineSpacing());

                doc.SaveAs(dst);
            }
        }
    }
}
