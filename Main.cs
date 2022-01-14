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
            string src = @"C:\Users\zhailiao123\Desktop\test.docx";
            string dst = @"C:\Users\zhailiao123\Desktop\test2.docx";
            //OP.WordprocessingDocument doc = OP.WordprocessingDocument.Open(filename, false);
            
            Document doc = new Document(src);
            Paragraph p = doc.Find("1").First();

            doc.Sections[0].Paragraphs.Remove(p);

            doc.SaveAs(dst);
            doc.Close();
            System.Diagnostics.Process.Start(dst);
        }
    }
}
