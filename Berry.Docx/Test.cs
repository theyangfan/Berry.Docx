using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using Berry.Docx;
using Berry.Docx.Documents;
using Berry.Docx.Field;

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
                ParagraphStyle style1 = new ParagraphStyle(doc);
                style1.Name = "样式1";
                style1.BaseStyle = ParagraphStyle.Default(doc);
                style1.AddToGallery = true;
                style1.IsCustom = true;
                
                doc.Styles.Add(style1);
                doc.SaveAs(dst);
            }
        }
    }
}
