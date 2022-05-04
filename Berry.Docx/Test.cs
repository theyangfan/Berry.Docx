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
            string src = @"C:\Users\Tomato\Desktop\test.docx";
            string dst = @"C:\Users\Tomato\Desktop\dst.docx";

            using (Document doc = new Document(src))
            {
                ParagraphStyle style = new ParagraphStyle(doc);
                style.BaseStyle = ParagraphStyle.Default(doc);
                style.Name = "样式a";
                style.AddToGallery = true;
                style.CharacterFormat.Bold = true;
                style.CharacterFormat.FontSize = 16;
                doc.Styles.Add(style);
                doc.SaveAs(dst);
            }
        }
    }
}
