using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using Berry.Docx.Field;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// The bookmark with a paired <see cref="BookmarkStart"/> and <see cref="BookmarkEnd"/>.
    /// </summary>
    public class Bookmark
    {
        #region Private Members
        private Document _doc;
        private readonly W.BookmarkStart _bookmarkStart;
        #endregion

        #region Constructors
        internal Bookmark(Document doc, W.BookmarkStart bookmarkStart)
        {
            _doc = doc;
            _bookmarkStart = bookmarkStart;
        }
        #endregion

        #region Public Properties

        /// <summary>
        /// Returns the unique identifier of the bookmark.
        /// </summary>
        public string Id => _bookmarkStart.Id;

        /// <summary>
        /// Returns the name of the bookmark.
        /// </summary>
        public string Name => _bookmarkStart.Name;

        /// <summary>
        /// Returns the start marker of the bookmark.
        /// </summary>
        public BookmarkStart Start => new BookmarkStart(_doc, _bookmarkStart);

        /// <summary>
        /// Returns the end marker of the bookmark.
        /// </summary>
        public BookmarkEnd End
        {
            get
            {
                W.Body body = _doc.Package.MainDocumentPart?.Document?.Body;
                if (body == null) return null;
                var bookmark = body.Descendants<W.BookmarkEnd>().Where(b => b.Id == _bookmarkStart.Id).FirstOrDefault();
                if(bookmark == null) return null;
                return new BookmarkEnd(_doc, bookmark);
            }
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Creates a new unique bookmark id.
        /// </summary>
        /// <param name="doc"></param>
        /// <returns>The new unique bookmark id.</returns>
        public static string CreateNewId(Document doc)
        {
            return IDGenerator.GenerateBookmarkId(doc);
        }
        #endregion
    }
}
