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
        public static void Main()
        {
            string src = @"C:\Users\tomato\Desktop\test.docx";
            using(Document doc = new Document(src, FileShare.ReadWrite))
            {
                var p = new Paragraph(doc);
                p.AppendText("图表");
                p.ChildItems.Add(new FieldChar(doc, FieldCharType.Begin));
                p.ChildItems.Add(new FieldCode(doc, "SEQ 图表 \\* ARABIC"));
                p.ChildItems.Add(new FieldChar(doc, FieldCharType.Separate));
                p.AppendText("1");
                p.ChildItems.Add(new FieldChar(doc, FieldCharType.End));
                
                doc.LastSection.ChildObjects.Add(p);
                doc.Save();
            }
        }
    }
}
