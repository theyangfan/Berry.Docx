using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// This class specifies that a non breaking hyphen character shall be 
    /// placed at the current location in the paragraph. A non breaking hyphen
    /// is the equivalent of Unicode character 002D (the hyphen-minus), however 
    /// it shall not be used as a line breaking character for the current line 
    /// of text when displaying this content.
    /// </summary>
    public class NoBreakHyphen : TextRange
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.NoBreakHyphen _hyphen;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new non breaking hyphen character.
        /// </summary>
        /// <param name="doc"></param>
        public NoBreakHyphen(Document doc) : this(doc, ParagraphItemGenerator.GenerateNoBreakHyphen())
        {
        }

        internal NoBreakHyphen(Document doc, W.NoBreakHyphen hyphen) : this(doc, hyphen.Parent as W.Run, hyphen)
        {
        }

        internal NoBreakHyphen(Document doc, W.Run ownerRun, W.NoBreakHyphen hyphen)
            : base(doc, ownerRun, hyphen)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _hyphen = hyphen;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the type of the current objetc.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.NoBreakHyphen;

        /// <summary>
        /// The "-" character.
        /// </summary>
        public override string Text
        {
            get => "-";
            set => throw new NotSupportedException("The Hyphen Character does not support modifying.");
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
            W.NoBreakHyphen hyphen = (W.NoBreakHyphen)_hyphen.CloneNode(true);
            run.RunProperties = _ownerRun.RunProperties?.CloneNode(true) as W.RunProperties; // copy format
            run.AppendChild(hyphen);
            return new NoBreakHyphen(_doc, run, hyphen);
        }
        #endregion
    }
}
