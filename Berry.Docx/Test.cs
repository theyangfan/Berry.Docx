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
            using(Document doc = new Document(src, FileShare.ReadWrite))
            {
                var p = doc.LastSection.Paragraphs[1];
                for(int i = 0; i < p.ChildItems.Count; i++)
                {
                    var item = p.ChildItems[i];
                    Console.WriteLine(item);
                    if(item is SimpleField)
                    {
                        p.ChildItems.InsertAt(new TextRange(doc, (item as SimpleField).Result), i);
                        p.ChildItems.Remove(item);
                    }
                    if (item is FieldChar || item is FieldCode)
                    {
                        p.ChildItems.Remove(item);
                        i--;
                    }
                }
                doc.Save();
            }
        }
    }
}
