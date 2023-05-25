using System;

namespace Berry.Docx.VisualModel
{
    public class Class1
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
