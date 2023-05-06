using System;
using System.Collections.Generic;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represents the start of a bookmark. This start marker is matched with the appropriately
    /// paired end marker by matching the value of the id from the associated <see cref="BookmarkEnd"/>.
    /// </summary>
    public class BookmarkStart : ParagraphItem
    {
        #region Private Members
        private readonly W.BookmarkStart _bookmarkStart;
        #endregion

        #region Constructor
        /// <summary>
        /// Creates a new BookmarkStart instance with the specifed id and name.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="id">A unique identifier for the bookmark.</param>
        /// <param name="name">The name of the bookmark. If multiple bookmarks
        /// in a document share the same name, the the first bookmark shall be maintained.</param>
        public BookmarkStart(Document doc, string id, string name) : this(doc, ParagraphItemGenerator.GenerateBookmarkStart(id, name))
        {

        }

        internal BookmarkStart(Document doc, W.BookmarkStart start) : base(doc, start)
        {
            _bookmarkStart = start;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the type of the current document object.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.BookmarkStart;

        /// <summary>
        /// Gets or sets the unique identifier for the bookmark.
        /// </summary>
        public string Id
        {
            get => _bookmarkStart.Id;
            set => _bookmarkStart.Id = value;
        }

        /// <summary>
        /// Gets or sets the name of the bookmark. If multiple bookmarks
        /// in a document share the same name, the the first bookmark
        /// shall be maintained.
        /// </summary>
        public string Name
        {
            get => _bookmarkStart.Name;
            set => _bookmarkStart.Name = value;
        }
        #endregion

    }
}
