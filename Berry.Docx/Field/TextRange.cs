using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;
using Berry.Docx.Collections;
using Berry.Docx.Formatting;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represent the text range.
    /// </summary>
    public class TextRange : ParagraphItem
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
        /// Gets the owner paragraph.
        /// </summary>
        public Paragraph OwnerParagraph
        {
            get
            {
                W.Paragraph paragraph = _run.Ancestors<W.Paragraph>().FirstOrDefault();
                if (paragraph != null)
                    return new Paragraph(_doc, paragraph);
                return null;
            }
        }
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

        #region Public Methods
        public CharacterStyle GetStyle()
        {
            if (_run?.RunProperties?.RunStyle != null)
            {
                W.Styles styles = _doc.Package.MainDocumentPart.StyleDefinitionsPart.Styles;
                string styleId = _run.RunProperties.RunStyle.Val.ToString();
                W.Style style =  styles.Elements<W.Style>().Where(s => s.StyleId == styleId).FirstOrDefault();
                if(style != null)
                    return new CharacterStyle(_doc, style);
            }
            return null;
        }

        public void ApplyStyle(string styleName)
        {
            if (_run == null || string.IsNullOrEmpty(styleName)) return;
            if (Style.NameToBuiltIn(styleName) != BuiltInStyle.None)
            {
                ApplyStyle(Style.NameToBuiltIn(styleName));
                return;
            }
            var style = _doc.Styles.FindByName(styleName, StyleType.Paragraph);
            if (style == null)
            {
                style = new ParagraphStyle(_doc, styleName);
                _doc.Styles.Add(style);
            }
            var linkedStyle = (style as ParagraphStyle).CreateLinkedStyle();
            if (_run.RunProperties == null)
                _run.RunProperties = new W.RunProperties();
            _run.RunProperties.RunStyle = new W.RunStyle() { Val = linkedStyle.StyleId };
        }

        public void ApplyStyle(BuiltInStyle bstyle)
        {
            if(_run == null) return;
            var style = ParagraphStyle.CreateBuiltInStyle(bstyle, _doc);
            if (style != null)
            {
                if (bstyle == BuiltInStyle.Normal)
                {
                    if (_run.RunProperties?.RunStyle != null)
                        _run.RunProperties.RunStyle = null;
                }
                else
                {
                    var linkedStyle = style.CreateLinkedStyle();
                    if (_run.RunProperties == null)
                        _run.RunProperties = new W.RunProperties();
                    _run.RunProperties.RunStyle = new W.RunStyle() { Val = linkedStyle.StyleId };
                }
            }
        }
        #endregion
    }
}
