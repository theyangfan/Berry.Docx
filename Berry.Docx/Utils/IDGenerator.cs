using System;
using System.Collections;
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
        public static int CUSTOM_NUM_ID = 1;
        public static int CUSTOM_ABSTRACT_NUM_ID = 0;
        public static int CUSTOM_BOOKMARK_ID = 0;
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

        public static int GenerateNumId(Document doc)
        {
            List<int> ids = new List<int>();
            Numbering numbering = doc.Package.MainDocumentPart.NumberingDefinitionsPart?.Numbering;
            if (numbering != null)
            {
                foreach(NumberingInstance num in numbering.Elements<NumberingInstance>())
                {
                    ids.Add(num.NumberID);
                }
            }
            while(ids.Contains(CUSTOM_NUM_ID)) CUSTOM_NUM_ID++;
            return CUSTOM_NUM_ID;
        }

        public static int GenerateAbstractNumId(Document doc)
        {
            List<int> ids = new List<int>();
            Numbering numbering = doc.Package.MainDocumentPart.NumberingDefinitionsPart?.Numbering;
            if (numbering != null)
            {
                foreach (AbstractNum num in numbering.Elements<AbstractNum>())
                {
                    ids.Add(num.AbstractNumberId);
                }
            }
            while (ids.Contains(CUSTOM_ABSTRACT_NUM_ID)) CUSTOM_ABSTRACT_NUM_ID++;
            return CUSTOM_ABSTRACT_NUM_ID;
        }

        public static string GenerateBookmarkId(Document doc)
        {
            List<int> ids = new List<int>();
            Body body = doc.Package.MainDocumentPart?.Document?.Body;
            if (body == null) return "0";
            foreach(var bookmark in body.Descendants<BookmarkStart>())
            {
                if (!string.IsNullOrEmpty(bookmark.Id?.Value))
                {
                    int.TryParse(bookmark.Id.Value, out int id);
                    ids.Add(id);
                }
            }
            while (ids.Contains(CUSTOM_BOOKMARK_ID)) ++CUSTOM_BOOKMARK_ID;
            return CUSTOM_BOOKMARK_ID.ToString();
        }
    }
}
