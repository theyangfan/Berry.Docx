using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Utils
{
    internal class RelationshipIdGenerator
    {
        public static string Generate(Document doc)
        {
            List<int> ids = new List<int>();
            foreach (var part in doc.Package.MainDocumentPart.Parts)
            {
                ids.Add(part.RelationshipId.Remove(0, 3).ToInt());
            }
            ids.Sort();
            return $"rId{ids.Last()+1}";
        }
    }
}
