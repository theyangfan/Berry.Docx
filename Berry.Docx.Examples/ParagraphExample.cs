using System;
using System.Collections.Generic;
using System.Text;
using Berry.Docx;
using Berry.Docx.Documents;
using Berry.Docx.Field;

namespace Berry.Docx.Examples
{
    public class ParagraphExample
    {
        public static void AddParagraph(Document doc)
        {
            // 1
            Paragraph p1 = new Paragraph(doc);
            p1.Text = "满纸荒唐言，一把辛酸泪。";
            doc.Sections[0].Paragraphs.Add(p1);

            // 2
            Paragraph p2 = doc.CreateParagraph();
            p2.Text = "都云作者痴，谁解其中味？";
            doc.Sections[0].ChildObjects.Add(p2);

            // 3
            Paragraph p3 = doc.Sections[0].AddParagraph();
            p3.Text = "字字看来皆是血，十年辛苦不寻常！";
        }

        public static void SetParagraphFormat(Document doc)
        {
            Paragraph paragraph = doc.Sections[0].Paragraphs[0];
            foreach(DocumentObject obj in paragraph.ChildObjects)
            {
                if(obj is TextRange)
                {
                    TextRange tr = obj as TextRange;
                    tr.CharacterFormat.FontNameEastAsia = "宋体";
                    tr.CharacterFormat.FontNameAscii = "Times New Roman";
                    tr.CharacterFormat.FontSize = 16;
                    tr.CharacterFormat.FontSizeCs = 16;
                }
            }
            paragraph.Format.Justification = JustificationType.Center;
            paragraph.Format.OutlineLevel = OutlineLevelType.Level1;
            paragraph.Format.LineSpacing = 24; // 2 lines
            paragraph.Format.LineSpacingRule = LineSpacingRule.Multiple;
        }
    }
}
