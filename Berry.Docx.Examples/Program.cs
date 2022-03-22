
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

namespace Berry.Docx.Examples
{
    public class Example
    {
        public static void Main()
        {
            string src = @"C:\Users\tomato\Desktop\test.docx";
            string dst = @"C:\Users\tomato\Desktop\dst.docx";

            Document doc = new Document(src);

            ParagraphExample.AddParagraph(doc);
            ParagraphExample.SetParagraphFormat(doc);

            // Save and close
            doc.SaveAs(dst);
            doc.Close();
            Console.WriteLine("----------------END----------------");
        }

    }
}
