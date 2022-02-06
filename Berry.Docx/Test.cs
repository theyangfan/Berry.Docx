using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Berry.Docx;
using Berry.Docx.Documents;
using Berry.Docx.Field;

using OP = DocumentFormat.OpenXml.Packaging;
using OW = DocumentFormat.OpenXml.Wordprocessing;

namespace Test
{
    public class Test
    {
        public static void Main() {
            string src = @"C:\Users\tomato\Desktop\test.docx";
            string dst = @"C:\Users\tomato\Desktop\test2.docx";
            //OP.WordprocessingDocument doc = OP.WordprocessingDocument.Open(filename, false);
            Document doc = new Document(src);

            Paragraph p1 = doc.CreateParagraph();
            p1.Text = "这是一个段落。";
            p1.CharacterFormat.FontEN = "微软雅黑";
            p1.CharacterFormat.FontSize = 14;
            p1.Format.Justification = JustificationType.Center;

            Table tbl1 = doc.CreateTable(3, 3);
            tbl1.Rows[0].Cells[1].Paragraphs[0].Text = "第1列";
            tbl1.Rows[0].Cells[2].Paragraphs[0].Text = "第2列";
            tbl1.Rows[1].Cells[0].Paragraphs[0].Text = "第1行";
            tbl1.Rows[2].Cells[0].Paragraphs[0].Text = "第2行";

            doc.Sections[0].Range.ChildObjects.Add(p1);
            doc.Sections[0].Range.ChildObjects.Add(tbl1);

            doc.Save();
            doc.Close();
            //System.Diagnostics.Process.Start(dst);
        } 
    }
}
