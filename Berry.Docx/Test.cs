using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Drawing;
using System.Diagnostics;

using Berry.Docx;
using Berry.Docx.Documents;
using Berry.Docx.Field;
using Berry.Docx.Formatting;

using P = DocumentFormat.OpenXml.Packaging;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Test
{
    internal class Test
    {
        public static void Main() {
            string src = @"C:\Users\Zhailiao123\Desktop\bugs\北京开放大学-原稿-702(校正).docx";
            string dst = @"C:\Users\Zhailiao123\Desktop\test\dst.docx";
            Stopwatch sw = Stopwatch.StartNew();
            sw.Start();
            // 打开指定文档
            using (Document doc = new Document(src))
            {
                // 打印所有段落的编号
                foreach (Paragraph p in doc.Paragraphs)
                {
                    if(!string.IsNullOrEmpty(p.ListText))
                    Console.WriteLine(p.ListText);
                }


                // doc.SaveAs(dst);
            }
            sw.Stop();
            Console.WriteLine(sw.Elapsed);
        }
    }
}
