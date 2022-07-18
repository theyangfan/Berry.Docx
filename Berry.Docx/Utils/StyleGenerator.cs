using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx
{
    internal class StyleGenerator
    {
        public static Style GenerateCharacterStyle(Document doc)
        {
            Style style = new Style();
            style.StyleId = IDGenerator.GenerateStyleID(doc);
            style.Type = StyleValues.Character;
            return style;
        }

        public static Style GenerateParagraphStyle(Document doc)
        {
            Style style = new Style();
            style.StyleId = IDGenerator.GenerateStyleID(doc);
            style.Type = StyleValues.Paragraph;
            return style;
        }

        public static Style GenerateTableStyle(Document doc)
        {
            Style style = new Style();
            style.StyleId = IDGenerator.GenerateStyleID(doc);
            style.Type = StyleValues.Table;
            return style;
        }

        public static Style GenerateNumberingStyle(Document doc)
        {
            Style style = new Style();
            style.StyleId = IDGenerator.GenerateStyleID(doc);
            style.Type = StyleValues.Numbering;
            return style;
        }
    }
}
