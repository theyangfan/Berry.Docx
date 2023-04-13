using System;
using System.Collections.Generic;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represent a revision of inserted range in the paragraph.
    /// </summary>
    public class InsertedRange : ParagraphItem
    {
        #region Priavate Members
        private readonly Document _doc;
        private readonly W.InsertedRun _ins;
        #endregion

        #region Constructors
        internal InsertedRange(Document doc, W.InsertedRun ins) : base(doc, ins)
        {
            _doc = doc;
            _ins = ins;
        }

        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the type of the current object.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.InsertedRange;

        /// <summary>
        /// Gets the inserted text.
        /// </summary>
        public string Text
        {
            get
            {
                StringBuilder text = new StringBuilder();
                foreach (var item in ChildObjects)
                {
                    if (item is TextRange)
                    {
                        text.Append(((TextRange)item).Text);
                    }
                    else if(item is Tab)
                    {
                        text.Append("\t");
                    }
                }
                return text.ToString();
            }
            set
            {
                _ins.RemoveAllChildren<W.Run>();
                TextRange tr = new TextRange(_doc);
                tr.Text = value;
                ChildObjects.Add(tr);
            }
        }
        #endregion
    }
}
