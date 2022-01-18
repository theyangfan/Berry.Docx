using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Berry.Docx;
using Berry.Docx.Documents;

using OP = DocumentFormat.OpenXml.Packaging;
using OW = DocumentFormat.OpenXml.Wordprocessing;

namespace Test
{
    public class Test
    {
        public static void Main() {
            string src = @"C:\Users\Tomato\Desktop\test.docx";
            string dst = @"C:\Users\Tomato\Desktop\test2.docx";
            //OP.WordprocessingDocument doc = OP.WordprocessingDocument.Open(filename, false);
            
            Document doc = new Document(src);

            Paragraph p = new Paragraph(doc) { Text = "这是2个段落" };
            Table table = new Table(doc, 10, 10);

            doc.Sections[0].Range.ChildObjects.Add(p);
            doc.Sections[0].Range.ChildObjects.Add(table);

            doc.SaveAs(dst);
            doc.Close();
            System.Diagnostics.Process.Start(dst);
        }
    }
}
