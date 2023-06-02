using System;
using System.Drawing;
using System.Windows;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Berry.Docx.Visual
{
    internal class Test
    {
        public static void Main()
        {
            string src = @"C:\Users\tomato\Desktop\test.docx";
            // 
            using(Berry.Docx.Document doc = new Berry.Docx.Document(src, System.IO.FileShare.ReadWrite))
            {
                Document visualDoc = new Document(doc);
                // get first page
                var page1 = visualDoc.Pages[0];
                Console.WriteLine(page1.Width);
                Console.WriteLine(page1.Height);
                Console.WriteLine(page1.Padding);
                // get first paragraph
                var paragraph1 = page1.Paragraphs[0];
                Console.WriteLine(paragraph1.Padding);
                // get first line
                var line1 = paragraph1.Lines[0];
                Console.WriteLine(line1.Height);
                Console.WriteLine(line1.HorizontalAlignment);
                // get first paragraph item
                var item1 = line1.ChildItems[1];
                Console.WriteLine(item1.Width);
                Console.WriteLine(item1.Height);
            }
        }

    }
}
