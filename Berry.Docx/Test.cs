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
using Berry.Docx.Formatting;

using P = DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Test
{
    internal class Test
    {
        public static void Main() {
            string src = @"C:\Users\Zhailiao123\Desktop\test\test.docx";
            string dst = @"C:\Users\Zhailiao123\Desktop\test\dst.docx";

            string str = " ㊻【页眉页脚结束2】";
            Regex rx = new Regex(@"^\s*\u0002?\s*([①-⑳]|[㉑-㉟]|[㊱-㊿])s*");
            Console.WriteLine(rx.IsMatch(str));
            Console.WriteLine(rx.Match(str).Value);
            return;
            using (Document doc = new Document(src))
            {
                Paragraph p = doc.LastSection.Paragraphs[1];
                foreach(DocumentItem item in p.ChildItems)
                {
                    Console.WriteLine(item.DocumentObjectType);
                }
                
                //doc.SaveAs(dst);
            }
        }
    }
}
