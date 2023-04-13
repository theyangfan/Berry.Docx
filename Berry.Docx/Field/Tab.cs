using System;
using System.Collections.Generic;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represent a tab character in the paragraph.
    /// </summary>
    public class Tab : TextRange
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.TabChar _tab;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new tab character.
        /// </summary>
        /// <param name="doc"></param>
        public Tab(Document doc) : this(doc, ParagraphItemGenerator.GenerateTab())
        {
        }

        internal Tab(Document doc, W.TabChar tab) : this(doc, tab.Parent as W.Run, tab)
        {
        }

        internal Tab(Document doc, W.Run ownerRun, W.TabChar tab)
            : base(doc, ownerRun, tab)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _tab = tab;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the type of the current objetc.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Tab;

        /// <summary>
        /// The "\t" character.
        /// </summary>
        public override string Text
        {
            get => "\t";
            set => throw new NotSupportedException("The Tab Character does not support modifying.");
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Creates a duplicate of the object.
        /// </summary>
        /// <returns>The cloned object.</returns>
        public override DocumentObject Clone()
        {
            W.Run run = new W.Run();
            W.TabChar tab = (W.TabChar)_tab.CloneNode(true);
            run.RunProperties = _ownerRun.RunProperties?.CloneNode(true) as W.RunProperties; // copy format
            run.AppendChild(tab);
            return new Tab(_doc, run, tab);
        }
        #endregion
    }
}
