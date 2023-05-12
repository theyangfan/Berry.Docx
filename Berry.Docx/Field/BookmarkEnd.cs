using System;
using System.Collections.Generic;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represents the end of a bookmark. This end marker is matched with the appropriately
    /// paired start marker by matching the value of the id from the associated <see cref="BookmarkStart"/>.
    /// </summary>
    public class BookmarkEnd : ParagraphItem
    {
        #region Private Members
        private readonly W.BookmarkEnd _bookmarkEnd;
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new BookmarkStart instance with the specifed id. 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="id">A unique identifier for the bookmark. The id should be the same with the previous <see cref="BookmarkStart"/>.</param>
        public BookmarkEnd(Document doc, string id) : this(doc, ParagraphItemGenerator.GenerateBookmarkEnd(id))
        {

        }

        internal BookmarkEnd(Document doc, W.BookmarkEnd end) : base(doc, end)
        {
            _bookmarkEnd = end;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the type of the current document object.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.BookmarkEnd;

        /// <summary>
        /// Gets or sets the unique identifier for the bookmark. The id should be the same with the previous <see cref="BookmarkStart"/>.
        /// </summary>
        public string Id
        {
            get => _bookmarkEnd.Id;
            set => _bookmarkEnd.Id = value;
        }
        #endregion
    }
}
