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
            Table tbl1 = doc.CreateTable(3, 3);// 3行3列

            tbl1.Rows[0].Cells[1].Paragraphs[0].Text = "第1列";
            tbl1.Rows[0].Cells[2].Paragraphs[0].Text = "第2列";
            tbl1.Rows[1].Cells[0].Paragraphs[0].Text = "第1行";
            tbl1.Rows[2].Cells[0].Paragraphs[0].Text = "第2行";

            doc.Sections[0].ChildObjects.Add(tbl1);
        }

        public static void InsertRowsAndColumns(Document doc)
        {
            Table tbl1 = doc.Sections[0].Tables[0];
            TableCell cell1 = tbl1.Rows[1].Cells[1];
            cell1.InsertRowAbove();
            cell1.InsertRowBelow();
            cell1.InsertColumnLeft();
            cell1.InsertColumnRight();
        }

    }
}
