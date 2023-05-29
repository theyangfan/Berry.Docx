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
            using(Berry.Docx.Document doc = new Berry.Docx.Document("example.docx"))
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
                // get first character
                var char1 = line1.Characters[0];
                Console.WriteLine(char1.Val);
                Console.WriteLine(char1.Width);
                Console.WriteLine(char1.Height);
            }
        }

    }
}
