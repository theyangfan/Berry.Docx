using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Berry.Docx.Documents;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Collections
{
    public class BookmarkCollection : IEnumerable<Bookmark>
    {
        private Document _doc;
        internal BookmarkCollection(Document doc)
        {
            _doc = doc;
        }

        public IEnumerator<Bookmark> GetEnumerator()
        {
            W.Body body = _doc.Package.MainDocumentPart?.Document?.Body;
            if (body == null) yield break;
            foreach (var bookmark in body.Descendants<W.BookmarkStart>())
            {
                if (!string.IsNullOrEmpty(bookmark.Id?.Value))
                {
                    yield return new Bookmark(_doc, bookmark);
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
