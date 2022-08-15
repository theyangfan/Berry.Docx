using System;
using System.Collections.Generic;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    public class Tab : ParagraphItem
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.TabChar _tab;
        #endregion

        #region Constructors
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
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Tab;

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
