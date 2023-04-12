using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// The FieldCode class specifies a field code within a complex field in the document.
    /// </summary>
    public class FieldCode : ParagraphItem
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.FieldCode _fieldCode;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new empty FieldCode.
        /// </summary>
        /// <param name="doc"></param>
        public FieldCode(Document doc) : this(doc, string.Empty) { }

        /// <summary>
        /// Initializes a new FieldCode with the specified code.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="code"></param>
        public FieldCode(Document doc, string code) : this(doc, ParagraphItemGenerator.GenerateFieldCode())
        {
            Code = code;
        }

        internal FieldCode(Document doc, W.FieldCode fieldCode) : this(doc, fieldCode.Parent as W.Run, fieldCode) { }

        internal FieldCode(Document doc, W.Run ownerRun, W.FieldCode fieldCode) : base(doc, ownerRun, fieldCode)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _fieldCode = fieldCode;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the type of the current object.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.FieldCode;

        /// <summary>
        /// Gets or sets the code string of the FieldCode.
        /// </summary>
        public string Code
        {
            get => _fieldCode.Text;
            set
            {
                _fieldCode.Text = value;
                if (Regex.IsMatch(value, @"\s"))
                {
                    _fieldCode.Space = O.SpaceProcessingModeValues.Preserve;
                }
                else
                {
                    _fieldCode.Space = null;
                }
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
            W.FieldCode fieldCode = (W.FieldCode)_fieldCode.CloneNode(true);
            run.RunProperties = _ownerRun.RunProperties?.CloneNode(true) as W.RunProperties; // copy format
            run.AppendChild(fieldCode);
            return new FieldCode(_doc, run, fieldCode);
        }
        #endregion
    }
}
