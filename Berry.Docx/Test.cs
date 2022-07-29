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
            string src = @"C:\Users\Zhailiao123\Desktop\bugs\14.大唐渭河热电厂热控专业检修工艺规程 389(校正).docx";
            //string src = @"C:\Users\Zhailiao123\Desktop\test\test.docx";
            string dst = @"C:\Users\Zhailiao123\Desktop\test\dst.docx";
            bool begin = false;
            using (Document doc = new Document(src))
            {
                foreach(Paragraph p in doc.Paragraphs)
                {
                    if (p.Text.Contains("电气线路的绝缘测试条件"))
                    {
                        begin = true;
                    }
                    if(begin) Console.WriteLine(p.Text);
                }
                //doc.SaveAs(dst);
            }
        }
    }
}
