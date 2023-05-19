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
            string src = @"C:\Users\zhailiao123\Desktop\docs\debug\test.docx";
            string dst = @"C:\Users\zhailiao123\Desktop\docs\debug\dst.docx";

            using (Document doc = new Document(src, FileShare.ReadWrite))
            {
                // 对文档最后一个段落进行标记
                Paragraph p1 = doc.LastSection.Paragraphs.Last();
                // 生成唯一的书签ID
                string bookmarkId = Bookmark.CreateNewId(doc);
                string bookmarkName = "_test" + bookmarkId;
                // 书签起始标签
                BookmarkStart bookmarkStart = new BookmarkStart(doc, bookmarkId, bookmarkName);
                // 书签结束标签
                BookmarkEnd bookmarkEnd = new BookmarkEnd(doc, bookmarkId);
                // 在段落开头插入起始标签
                p1.ChildItems.InsertAt(bookmarkStart, 0);
                // 在段落末尾添加结束标签
                p1.ChildItems.Add(bookmarkEnd);

                // 在文档第一个段落添加引用上面书签的域
                Paragraph p2 = doc.LastSection.Paragraphs[0];

                // 引用页码
                FieldChar fieldBegin1 = new FieldChar(doc, FieldCharType.Begin);
                FieldCode fieldCode1 = new FieldCode(doc, $" PAGEREF {bookmarkName} \\h ");
                FieldChar fieldSeparate1 = new FieldChar(doc, FieldCharType.Separate);
                TextRange result1 = new TextRange(doc, "1");
                FieldChar fieldEnd1 = new FieldChar(doc, FieldCharType.End);

                p2.ChildItems.Add(fieldBegin1);
                p2.ChildItems.Add(fieldCode1);
                p2.ChildItems.Add(fieldSeparate1);
                p2.ChildItems.Add(result1);
                p2.ChildItems.Add(fieldEnd1);

                // 保存
                doc.SaveAs(dst);
            }
        }
    }
}
