using System;

namespace Berry.Docx.Visual
{
    public class Test
    {
        public static void Main()
        {
            string src = @"C:\Users\tomato\Desktop\test.docx";
            using(var doc = new Berry.Docx.Document(src, System.IO.FileShare.ReadWrite))
            {
                Document document = new Document(doc);
                Console.WriteLine(document.Pages.Count);
            }
        }
    }
}
