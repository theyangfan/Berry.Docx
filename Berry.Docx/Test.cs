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
            
            Document doc = new Document(src);
            Paragraph p = doc.Sections[0].Paragraphs[0];
            TextRange tr = p.ChildObjects[0] as TextRange;

            tr.CharacterFormat.Position = 1.2f;
            p.CharacterFormat.Position = 1.3f;
            p.Style.CharacterFormat.Position = -1.8f;

            Console.WriteLine(tr.CharacterFormat.Position);
            Console.WriteLine(p.CharacterFormat.Position);
            Console.WriteLine(p.Style.CharacterFormat.Position);

            doc.SaveAs(dst);
            doc.Close();

            //System.Diagnostics.Process.Start(dst);
        }

    }
}
