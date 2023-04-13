using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// The SimpleField class specifies the presence of a simple field at the current location in the document.
    /// </summary>
    public class SimpleField : ParagraphItem
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.SimpleField _fldSimple;
        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a SimpleField instance with the specified code and result.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="code">The field code.</param>
        /// <param name="result">The field result</param>
        public SimpleField(Document doc, string code, string result)
            : this(doc, ParagraphItemGenerator.GenerateSimpleField())
        {
            Code = code;
            Result = result;
        }

        internal SimpleField(Document doc, W.SimpleField fldSimple) : base(doc, fldSimple)
        {
            _doc = doc;
            _fldSimple = fldSimple;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the type of the current object.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.SimpleField;

        /// <summary>
        /// Gets or sets the field code.
        /// </summary>
        public string Code
        {
            get
            {
                return _fldSimple.Instruction;
            }
            set
            {
                _fldSimple.Instruction = value;
            }
        }

        /// <summary>
        /// Gets or sets the result of the field code.
        /// </summary>
        public string Result
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                foreach(var item in ChildObjects.OfType<TextRange>())
                {
                    sb.Append(item.Text);
                }
                return sb.ToString();
            }
            set
            {
                TextRange tr = ChildObjects.OfType<TextRange>().FirstOrDefault()?.Clone() as TextRange;
                if (tr == null) tr = new TextRange(_doc);
                ChildObjects.RemoveAll<TextRange>();
                ChildObjects.Add(tr);
                tr.Text = value;
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
            W.SimpleField fldSimple = (W.SimpleField)_fldSimple.CloneNode(true);
            return new SimpleField(_doc, fldSimple);
        }
        #endregion
    }
}
