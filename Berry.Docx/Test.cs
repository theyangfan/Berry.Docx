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
                //TextRange tr = p.ChildItems[2] as TextRange;
                //p.Format.SetLeftIndent(0, IndentationUnit.Point);
                //p.Format.SetRightIndent(1.5f, IndentationUnit.Character);
                //Console.WriteLine(p.Format.GetLeftIndent());
                //Console.WriteLine(p.Format.GetRightIndent());
                Console.WriteLine(p.Format.GetSpecialIndentation());
                //Console.WriteLine(p.GetStyle().ParagraphFormat.GetLeftIndent());
                doc.SaveAs(dst);
            }
        }
    }
}
