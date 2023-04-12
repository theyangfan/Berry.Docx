using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx
{
    internal class ParagraphItemGenerator
    {
        public static Break GenerateBreak()
        {
            Run run = new Run();
            Break br = new Break();
            run.AddChild(br);
            return br;
        }

        public static TabChar GenerateTab()
        {
            Run run = new Run();
            TabChar tab = new TabChar();
            run.AddChild(tab);
            return tab;
        }

        public static FieldChar GenerateFieldChar()
        {
            Run run = new Run();
            FieldChar field = new FieldChar();
            run.AddChild(field);
            return field;
        }

        public static FieldCode GenerateFieldCode()
        {
            Run run = new Run();
            FieldCode field = new FieldCode();
            run.AddChild(field);
            return field;
        }

        public static SimpleField GenerateSimpleField()
        {
            SimpleField field = new SimpleField();
            return field;
        }
    }
}
