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
            string filename = @"C:\Users\zhailiao123\Desktop\test.docx";
            OP.WordprocessingDocument doc = OP.WordprocessingDocument.Open(filename, false);
            doc.MainDocumentPart.Document.Body

            
            Document doc = new Document(filename);
            foreach(Paragraph p in doc.Sections[1].Paragraphs)
                Console.WriteLine(p.Text);

            //doc.Save();
            doc.Close();
            
        }
    }
}
