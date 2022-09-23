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
            string dst = @"C:\Users\Zhailiao123\Desktop\test\test.png";

            using (Document doc = new Document(src, FileShare.ReadWrite))
            {
                Paragraph p = doc.Paragraphs.Last();
                foreach(var obj in p.ChildItems)
                {
                    Console.WriteLine(obj);
                    if(obj is Picture)
                    {
                        Picture pic = obj as Picture;
                        using(var stream = pic.Stream)
                        {
#if NET40_OR_GREATER
                            Image image = Image.FromStream(stream);
                            image.Save(dst);
#endif
                        }
                    }
                }
                //doc.SaveAs(dst);
            }
        }
    }
}
