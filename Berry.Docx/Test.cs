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
                TextRange tr = p.ChildItems[2] as TextRange;

                tr.CharacterFormat.TextColor = Color.Blue;
                tr.CharacterFormat.SubSuperScript = SubSuperScript.None;
                tr.CharacterFormat.UnderlineStyle = UnderlineStyle.None;

                Console.WriteLine(tr.CharacterFormat.SubSuperScript);
                Console.WriteLine(tr.CharacterFormat.UnderlineStyle);
                Console.WriteLine(tr.CharacterFormat.TextColor);

                doc.SaveAs(dst);
            }
        }
    }
}
