using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
        public static void Main()
        {
            string src = @"C:\Users\tomato\Desktop\test.docx";
            string dst = @"C:\Users\tomato\Desktop\dst.docx";
#if NET6_0
            Console.WriteLine(Convert.ToHexString(System.Text.Encoding.Unicode.GetBytes("我")));
#endif
            using (Document doc = new Document(src, FileShare.ReadWrite))
            {
                var paragraph = doc.LastSection.Paragraphs[0];
                foreach(var tr in paragraph.ChildItems.OfType<TextRange>())
                {
                    tr.CharacterFormat.FontNameComplexScript = "楷体";
                    tr.CharacterFormat.FontNameAscii = "Times New Roman";
                    tr.CharacterFormat.FontNameEastAsia = "微软雅黑";
                    //tr.CharacterFormat.FontTypeHint = FontContentType.EastAsia;
                    tr.CharacterFormat.UseComplexScript = true;
                    //tr.CharacterFormat.RightToLeft = true;
                }

                // 保存
                doc.SaveAs(dst);
            }
        }
    }
}
