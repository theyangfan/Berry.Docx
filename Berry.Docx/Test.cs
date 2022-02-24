﻿using System;
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
            string src = @"C:\Users\zhailiao123\Desktop\test.docx";
            string dst = @"C:\Users\tomato\Desktop\test2.docx";
            
            Document doc = new Document(src);
            Paragraph p = doc.Sections[0].Paragraphs[0];

            p.AppendComment("test", "5");
            doc.Save();
            doc.Close();

            //System.Diagnostics.Process.Start(dst);
        }
    }
}
