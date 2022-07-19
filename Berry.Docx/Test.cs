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

            using (Document doc = new Document(src))
            {
                Paragraph p = doc.LastSection.Paragraphs[0];
                TableStyle style = doc.Styles.FindByName("样式1", StyleType.Table) as TableStyle;
                if(style != null)
                {
                    style.FirstRow.ParagraphFormat.Justification = JustificationType.Both;
                    Console.WriteLine(style.FirstRow.ParagraphFormat.Justification);
                    Console.WriteLine(style.FirstRow.ParagraphFormat.OutlineLevel);
                }
                
                doc.SaveAs(dst);
            }
        }
    }
}
