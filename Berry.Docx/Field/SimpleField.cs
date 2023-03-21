using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represent a simple filed code.
    /// </summary>
    public class SimpleField : ParagraphItem
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.SimpleField _fldSimple;
        #endregion

        #region Constructor
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
                foreach (var item in ChildObjects.OfType<TextRange>())
                {
                    item.Remove();
                }
                ChildObjects.Add(tr);
                tr.Text = value;
            }
        }
        #endregion
    }
}
