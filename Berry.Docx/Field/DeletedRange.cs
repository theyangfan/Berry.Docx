using System;
using System.Collections.Generic;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represent a revision of deleted range in the paragraph.
    /// </summary>
    public class DeletedRange : ParagraphItem
    {
        #region Priavate Members
        private readonly W.DeletedRun _del;
        #endregion

        #region Constructors
        internal DeletedRange(Document doc, W.DeletedRun del) : base(doc, del)
        {
            _del = del;
        }

        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the type of the current object.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.DeletedRange;

        /// <summary>
        /// Gets the deleted text.
        /// </summary>
        public string Text
        {
            get
            {
                StringBuilder text = new StringBuilder();
                foreach (var item in ChildObjects)
                {
                    if (item is DeletedTextRange)
                    {
                        text.Append(((DeletedTextRange)item).Text);
                    }
                }
                return text.ToString();
            }
        }
        #endregion
    }
}
