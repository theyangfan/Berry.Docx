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

            }
        }

    }
}
