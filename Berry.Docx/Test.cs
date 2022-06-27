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
                TextRange tr = p.ChildItems[0] as TextRange;

                p.Format.Borders.Top.Style = BorderStyle.Nil;
                p.Format.Borders.Bottom.Style = BorderStyle.Nil;
                p.Format.Borders.Left.Style = BorderStyle.None;
                p.Format.Borders.Right.Style = BorderStyle.None;

                //tr.CharacterFormat.Border.Color = Color.Blue;
                //tr.CharacterFormat.Border.Style = BorderStyle.Inset;

                //Console.WriteLine(tr.CharacterFormat.Border.Color);
                Console.WriteLine(p.Format.Borders.Top.Style);
                Console.WriteLine(p.Format.Borders.Top.Color);
                Console.WriteLine(p.Format.Borders.Top.Width);
                

                doc.SaveAs(dst);
            }
        }
    }
}
