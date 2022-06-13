using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx
{
    internal class IDGenerator
    {
        public static int CUSTOM_STYLE_ID = 1;
        public static string GenerateRelationshipID(Document doc)
        {
            List<int> ids = new List<int>();
            foreach (var part in doc.Package.MainDocumentPart.Parts)
            {
                ids.Add(part.RelationshipId.Remove(0, 3).ToInt());
            }
            ids.Sort();
            return $"rId{ids.Last()+1}";
        }

        public static string GenerateStyleID(Document doc)
        {
            List<string> ids = new List<string>();
            foreach(Style style in doc.Package.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>())
            {
                ids.Add(style.StyleId);
            }
            do
            {
                CUSTOM_STYLE_ID++;
            } while (ids.Contains(CUSTOM_STYLE_ID.ToString()));
            return CUSTOM_STYLE_ID.ToString();
        }
    }
}
