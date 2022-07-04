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
                ListStyle style = ListStyle.Create(doc, BuiltInListStyle.Style1);
                style.Name = "样式1";
                style.Levels[0].Pattern = "第1章";
                style.Levels[0].NumberStyle = ListNumberStyle.Decimal;
                style.Levels[0].SuffixCharacter = LevelSuffixCharacter.Space;
                style.Levels[0].NumberPosition = 0.75f / 2.54f * 72;
                style.Levels[0].TextIndentation = 1.75f / 2.54f * 72;
                style.Levels[0].CharacterFormat.FontSize = 20;
                style.Levels[0].CharacterFormat.FontNameEastAsia = "微软雅黑";
                doc.ListStyles.Add(style);

                p.ListFormat.ApplyStyle("样式1", 1);

                doc.SaveAs(dst);
            }
        }
    }
}
