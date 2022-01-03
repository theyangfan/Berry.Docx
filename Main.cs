using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Berry.Docx;
using Berry.Docx.Documents;

namespace Test
{
    public class Test
    {
        public static void Main() {
            
            Document doc = new Document(@"C:\Users\Levi\Desktop\test.docx");
            
            Paragraph p = doc.Paragraphs[0];
            //字体
            p.CharacterFormat.FontCN = "微软雅黑";
            // 字号
            p.CharacterFormat.FontSize = 16;
            // 加粗
            p.CharacterFormat.Bold = true;
            // 对齐方式
            p.Format.Justification = JustificationType.Left;
            // 大纲级别
            p.Format.OutlineLevel = OutlineLevelType.Level1;
            // 左侧缩进
            p.Format.LeftCharsIndent = 2;
            // 段前行距
            p.Format.BeforeLinesSpacing = 1;
            //行距
            p.Format.LineSpacing = 12;
            p.Format.LineSpacingRule = LineSpacingRule.Multiple;
            
            doc.Save();
            doc.Close();
        }
    }
}
