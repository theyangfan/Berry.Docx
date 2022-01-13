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
            string filename = @"C:\Users\zhailiao123\Desktop\test7.docx";
            //OP.WordprocessingDocument doc = OP.WordprocessingDocument.Open(filename, false);
            
            Document doc = new Document(filename);

            Paragraph p = doc.Find("测试2").First();

            doc.Save();
            doc.Close();
            System.Diagnostics.Process.Start(filename);
        }
    }
}
