using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;
using Berry.Docx.Collections;
using Berry.Docx.Utils;
using Berry.Docx.Formatting;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represent the text range.
    /// </summary>
    public class TextRange : DocumentElement
    {
        #region Private Members
        private Document _doc;
        private W.Run _run;
        private W.Text _text;
        private CharacterFormat _cFormat;
        #endregion

        #region Constructors
        /// <summary>
        /// The TextRange constructor.
        /// </summary>
        /// <param name="doc">The owner document.</param>
        public TextRange(Document doc) : this(doc, RunGenerator.Generate(""))
        {
        }
        internal TextRange(Document doc, W.Run run) : base(doc, run)
        {
            _doc = doc;
            _run = run;
            _text = run.Elements<W.Text>().FirstOrDefault();
            _cFormat = new CharacterFormat(doc, run);
        }
        #endregion

        #region Public Properties

        /// <summary>
        /// The DocumentObject type.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.TextRange;

        /// <summary>
        /// The text.
        /// </summary>
        public string Text
        {
            get
            {
                if (_text != null)
                    return _text.Text;
                return string.Empty;
            }
            set
            {
                if(_text == null)
                {
                    _text = new W.Text();
                    _run.AddChild(_text);
                }
                _text.Text = value;
            }
        }

        /// <summary>
        /// The character format.
        /// </summary>
        public CharacterFormat CharacterFormat => _cFormat;
        #endregion
    }
}
