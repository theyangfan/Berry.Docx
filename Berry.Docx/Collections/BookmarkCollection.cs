using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Berry.Docx.Documents;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// Represents a colection of bookmarks.
    /// </summary>
    public class BookmarkCollection : IEnumerable<Bookmark>
    {
        #region Private Members
        private Document _doc;
        #endregion

        #region Constructors
        internal BookmarkCollection(Document doc)
        {
            _doc = doc;
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
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
        #endregion

    }
}
