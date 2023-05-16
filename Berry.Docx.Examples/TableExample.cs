using System;
using System.Collections.Generic;
using System.Text;
using Berry.Docx;
using Berry.Docx.Documents;
using Berry.Docx.Field;

namespace Berry.Docx.Examples
{
    public class TableExample
    {
        public static void AddTable(Document doc)
        {
            // 创建一个3行3列的表格
            Table tbl1 = doc.CreateTable(3, 3);

            // 设置单元格内容（每个单元格默认有一个段落）
            tbl1.Rows[0].Cells[1].Paragraphs[0].Text = "第1列";
            tbl1.Rows[0].Cells[2].Paragraphs[0].Text = "第2列";
            tbl1.Rows[1].Cells[0].Paragraphs[0].Text = "第1行";
            tbl1.Rows[2].Cells[0].Paragraphs[0].Text = "第2行";

            // 添加到文档末尾
            doc.LastSection.ChildObjects.Add(tbl1);

            // 在单元格周围插入行和列
            TableCell cell1 = tbl1.Rows[1].Cells[1];
            cell1.Paragraphs[0].Text = "在我周围插入行和列";
            cell1.InsertRowAbove();
            cell1.InsertRowBelow();
            cell1.InsertColumnLeft();
            cell1.InsertColumnRight();

            // 根据窗口自动调整表格宽度
            tbl1.AutoFit(AutoFitMethod.AutoFitWindow);
        }

    }
}
