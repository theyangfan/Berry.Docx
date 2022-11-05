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
        public static void Main() {
            string src = @"C:\Users\Zhailiao123\Desktop\test\test.docx";
            string dst = @"C:\Users\Zhailiao123\Desktop\test\dst.docx";
            string file = @"C:\Users\Zhailiao123\Desktop\test\test1.jpg";

            using (Document doc = new Document(src, FileShare.ReadWrite))
            {
                Table tbl = doc.Tables[0];
                TableCell cell = tbl[1][0];
                Console.WriteLine(cell.GetCellWidth().Val);
                Console.WriteLine(cell.GetCellWidth().Type);
                
                //doc.SaveAs(dst);
            }
        }
    }
}
