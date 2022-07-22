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
                //Paragraph p = doc.LastSection.Paragraphs[0];
                Table table = doc.LastSection.Tables[0];
                TableStyle style = new TableStyle(doc, "我的样式");

                style.WholeTable.VerticalCellAlignment = TableCellVerticalAlignment.Top;
                style.FirstRow.VerticalCellAlignment = TableCellVerticalAlignment.Bottom;

                style.FirstRow.CharacterFormat.Bold = true;
                style.WholeTable.ParagraphFormat.Justification = JustificationType.Center;

                style.FirstRow.Borders.Top.Style = BorderStyle.Single;
                style.FirstRow.Borders.Top.Color = Color.Black;
                style.FirstRow.Borders.Top.Width = 1.5f;

                style.FirstRow.Borders.Bottom.Style = BorderStyle.Single;
                style.FirstRow.Borders.Bottom.Color = Color.Black;
                style.FirstRow.Borders.Bottom.Width = 0.5f;

                style.LastRow.Borders.Bottom.Style = BorderStyle.Single;
                style.LastRow.Borders.Bottom.Color = Color.Black;
                style.LastRow.Borders.Bottom.Width = 1.5f;

                doc.Styles.Add(style);
                table.ApplyStyle("我的样式");
                doc.SaveAs(dst);
            }
        }
    }
}
