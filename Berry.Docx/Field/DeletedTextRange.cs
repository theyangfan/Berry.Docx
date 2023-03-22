using System;
using System.Collections.Generic;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represent a revision of deleted text in the paragraph.
    /// </summary>
    public class DeletedTextRange : ParagraphItem
    {
        #region Private Members
        private readonly W.DeletedText _text;
        #endregion

        #region Constructor
        internal DeletedTextRange(Document doc, W.Run ownerRun) : base(doc, ownerRun, ownerRun.GetFirstChild<W.DeletedText>())
        {
            _text = ownerRun.GetFirstChild<W.DeletedText>();
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the type of the current object.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.DeletedTextRange;

        /// <summary>
        /// Gets the deleted text.
        /// </summary>
        public string Text
        {
            get
            {
                return _text?.Text;
            }
        }
        #endregion
    }
}
