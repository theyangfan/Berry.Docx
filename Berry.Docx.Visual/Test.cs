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
            string src = @"C:\Users\zhailiao123\Desktop\docs\debug\test.docx";
            using(var doc = new Berry.Docx.Document(src, System.IO.FileShare.ReadWrite))
            {
                Document document = new Document(doc);
                Console.WriteLine(document.Pages.Count);
            }
        }

    }
}
