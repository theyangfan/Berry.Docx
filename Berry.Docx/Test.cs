using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

            using (Document doc = new Document(src, FileShare.ReadWrite))
            {
                //var paragraph = doc.LastSection.Paragraphs[0];
                var table = doc.LastSection.Tables[0];
                var c1 = table[0][0];
                var c2 = table[1][0];

                c1.Borders.Bottom.Style = BorderStyle.Single;
                c2.Borders.Top.Style = BorderStyle.Single;
                c1.Borders.Bottom.Width = 4;
                c1.Borders.Bottom.Color = Color.Red;
                c2.Borders.Top.Width = 1;
                c2.Borders.Top.Color = Color.Yellow;
                // 保存
                doc.SaveAs(dst);
            }
        }
    }
}
