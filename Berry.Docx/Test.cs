using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
            string src = @"C:\Users\tomato\Desktop\test.docx";
            string dst = @"C:\Users\tomato\Desktop\dst.docx";

            Document doc = new Document("example.docx");
            // Create a new paragraph
            Paragraph p1 = doc.CreateParagraph();
            p1.Text = "This is a paragraph.";
            p1.CharacterFormat.FontNameEastAsia = "Times New Roman";
            p1.CharacterFormat.FontSize = 14;
            p1.Format.Justification = JustificationType.Center;
            // Create a new table
            Table tbl1 = doc.CreateTable(3, 3);
            tbl1.Rows[0].Cells[1].Paragraphs[0].Text = "1st Column";
            tbl1.Rows[0].Cells[2].Paragraphs[0].Text = "2nd Column";
            tbl1.Rows[1].Cells[0].Paragraphs[0].Text = "1st Row";
            tbl1.Rows[2].Cells[0].Paragraphs[0].Text = "2nd Row";
            // Add to the document
            doc.Sections[0].ChildObjects.Add(p1);
            doc.Sections[0].ChildObjects.Add(tbl1);
            // Save and close
            doc.Save();
            doc.Close();

            //System.Diagnostics.Process.Start(dst);
        }

    }
}
