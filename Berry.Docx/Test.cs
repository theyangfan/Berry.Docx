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
            string src = @"C:\Users\Zhailiao123\Desktop\test.docx";
            string dst = @"C:\Users\Zhailiao123\Desktop\test2.docx";
            //OP.WordprocessingDocument doc = OP.WordprocessingDocument.Open(filename, false);
            
            Document doc = new Document(src);
            Table table = doc.Sections[0].Tables[0];
            Paragraph p = doc.CreateParagraph();
            p.Text = "new";

            doc.SaveAs(dst);
            doc.Close();
            System.Diagnostics.Process.Start(dst);
        }
    }
}
