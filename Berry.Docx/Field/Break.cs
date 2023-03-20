using System;
using System.Collections.Generic;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represent a break charater in the paragraph.
    /// </summary>
    public class Break : ParagraphItem
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.Break _break;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new break with the specified type.
        /// </summary>
        /// <param name="doc">The owner document.</param>
        /// <param name="type">The break type.</param>
        public Break(Document doc, BreakType type) : this(doc, ParagraphItemGenerator.GenerateBreak())
        {
            Type = type;
        }

        internal Break(Document doc, W.Break br) : this(doc, br.Parent as W.Run, br) { }

        internal Break(Document doc, W.Run ownerRun, W.Break br)
            : base(doc, ownerRun, br)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _break = br;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the type of the current object.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Break;

        /// <summary>
        /// 
        /// </summary>
        public BreakType Type
        {
            get
            {
                if (_break.Type == null) return BreakType.TextWrapping;
                return _break.Type.Value.Convert<BreakType>();
            }
            set
            {
                if (value == BreakType.TextWrapping)
                    _break.Type = null;
                else
                    _break.Type = value.Convert<W.BreakValues>();
            }
        }

        public BreakTextRestartLocation Clear
        {
            get
            {
                if(Type != BreakType.TextWrapping) return BreakTextRestartLocation.None;
                if (_break.Clear == null) return BreakTextRestartLocation.None;
                return _break.Clear.Value.Convert<BreakTextRestartLocation>();
            }
            set
            {
                if(value == BreakTextRestartLocation.None)
                    _break.Clear = null;
                else
                    _break.Clear = value.Convert<W.BreakTextRestartLocationValues>();
            }
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
            W.Break br = (W.Break)_break.CloneNode(true);
            run.RunProperties = _ownerRun.RunProperties?.CloneNode(true) as W.RunProperties; // copy format
            run.AppendChild(br);
            return new Break(_doc, run, br);
        }
        #endregion
    }
}
