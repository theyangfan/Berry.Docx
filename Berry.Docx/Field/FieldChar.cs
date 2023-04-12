using System;
using System.Collections.Generic;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// The FieldChar class specifies the presence of a complex field character at the current location in paragraph.
    /// A complex field character is a special character which delimits the start and end of a complex field or separates 
    /// its field codes from its current field result.
    /// </summary>
    public class FieldChar : ParagraphItem
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.FieldChar _fieldChar;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new FieldChar with the specified type.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="type"></param>
        public FieldChar(Document doc, FieldCharType type) : this(doc, ParagraphItemGenerator.GenerateFieldChar())
        {
            Type = type;
        }

        internal FieldChar(Document doc, W.FieldChar fieldChar) : this(doc, fieldChar.Parent as W.Run, fieldChar) { }

        internal FieldChar(Document doc, W.Run ownerRun, W.FieldChar fieldChar) : base(doc, ownerRun, fieldChar)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _fieldChar = fieldChar;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the type of the current object.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.FieldChar;

        /// <summary>
        /// Gets or sets the type of the FieldChar.
        /// </summary>
        public FieldCharType Type
        {
            get
            {
                if (_fieldChar.FieldCharType == null) return FieldCharType.Begin;
                return _fieldChar.FieldCharType.Value.Convert<FieldCharType>();
            }
            set
            {
                _fieldChar.FieldCharType = value.Convert<W.FieldCharValues>();
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
            W.FieldChar fieldChar = (W.FieldChar)_fieldChar.CloneNode(true);
            run.RunProperties = _ownerRun.RunProperties?.CloneNode(true) as W.RunProperties; // copy format
            run.AppendChild(fieldChar);
            return new FieldChar(_doc, run, fieldChar);
        }
        #endregion
    }
}
