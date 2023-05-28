using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text.RegularExpressions;

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
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private W.Text _text;
        private O.OpenXmlElement _element;
        private CharacterFormat _cFormat;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new empty TextRange.
        /// </summary>
        /// <param name="doc">The owner document.</param>
        public TextRange(Document doc) : this(doc, string.Empty)
        {
        }
        /// <summary>
        /// Initializes a new TextRange with the specified text.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="text"></param>
        public TextRange(Document doc, string text) : this(doc, ParagraphItemGenerator.GenerateTextRange(text))
        {
            Text = text;
        }

        internal TextRange(Document doc, W.Text text) : this(doc, text.Parent as W.Run, text)
        {
        }

        internal TextRange(Document doc, W.Run ownerRun, W.Text text) : base(doc, ownerRun, text)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _text = text;
            _cFormat = new CharacterFormat(doc, ownerRun);
        }

        internal TextRange(Document doc, W.Run ownerRun, O.OpenXmlElement element) : base(doc, ownerRun, element)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _element = element;
            _cFormat = new CharacterFormat(doc, ownerRun);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the DocumentObject type.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.TextRange;

        /// <summary>
        /// Gets or sets the text.
        /// </summary>
        public virtual string Text
        {
            get
            {
                if (_text != null)
                    return _text.Text;
                return string.Empty;
            }
            set
            {
                if (string.IsNullOrEmpty(value) && _ownerRun.Parent != null)
                {
                    this.Remove();
                    return;
                }
                if (_text == null)
                {
                    _text = new W.Text();
                    _ownerRun.AddChild(_text);
                }
                _text.Text = value;
                if(Regex.IsMatch(value, @"\s"))
                {
                    _text.Space = O.SpaceProcessingModeValues.Preserve;
                }
                else
                {
                    _text.Space = null;
                }
            }
        }

        /// <summary>
        /// Gets the character format.
        /// </summary>
        public override CharacterFormat CharacterFormat => _cFormat;

        /// <summary>
        /// Gets the characters of the current text range.
        /// </summary>
        public IEnumerable<Character> Characters
        {
            get
            {
                foreach(var c in Text)
                {
                    yield return new Character(c, _cFormat);
                }
            }
        }

        #endregion

        #region Public Methods
        /// <summary>
        /// Gets the character style.
        /// </summary>
        /// <returns></returns>
        public CharacterStyle GetStyle()
        {
            if (_ownerRun?.RunProperties?.RunStyle != null)
            {
                W.Styles styles = _doc.Package.MainDocumentPart.StyleDefinitionsPart.Styles;
                string styleId = _ownerRun.RunProperties.RunStyle.Val.ToString();
                W.Style style =  styles.Elements<W.Style>().Where(s => s.StyleId == styleId).FirstOrDefault();
                if(style != null)
                    return new CharacterStyle(_doc, style);
            }
            return null;
        }

        /// <summary>
        /// Applies the character style with the specified name. 
        /// If the specified style not exist, a new style with the specified name will be created.
        /// </summary>
        /// <param name="styleName">The style name.</param>
        public void ApplyStyle(string styleName)
        {
            if (_ownerRun == null || string.IsNullOrEmpty(styleName)) return;
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
            if (_ownerRun.RunProperties == null)
                _ownerRun.RunProperties = new W.RunProperties();
            _ownerRun.RunProperties.RunStyle = new W.RunStyle() { Val = linkedStyle.StyleId };
        }

        /// <summary>
        /// Applies the built-in style.
        /// </summary>
        /// <param name="bstyle">The built-in style.</param>
        public void ApplyStyle(BuiltInStyle bstyle)
        {
            if(_ownerRun == null) return;
            var style = ParagraphStyle.CreateBuiltInStyle(bstyle, _doc);
            if (style != null)
            {
                if (bstyle == BuiltInStyle.Normal)
                {
                    if (_ownerRun.RunProperties?.RunStyle != null)
                        _ownerRun.RunProperties.RunStyle = null;
                }
                else
                {
                    var linkedStyle = style.CreateLinkedStyle();
                    if (_ownerRun.RunProperties == null)
                        _ownerRun.RunProperties = new W.RunProperties();
                    _ownerRun.RunProperties.RunStyle = new W.RunStyle() { Val = linkedStyle.StyleId };
                }
            }
        }

        /// <summary>
        /// Creates a duplicate of the object.
        /// </summary>
        /// <returns>The cloned object.</returns>
        public override DocumentObject Clone()
        {
            W.Run run = new W.Run();
            if(_text != null)
            {
                W.Text text = (W.Text)_text.CloneNode(true);
                run.RunProperties = _ownerRun.RunProperties?.CloneNode(true) as W.RunProperties; // copy format
                run.AppendChild(text);
                return new TextRange(_doc, run, text);
            }
            else
            {
                var ele = _element.CloneNode(true);
                run.RunProperties = _ownerRun.RunProperties?.CloneNode(true) as W.RunProperties; // copy format
                run.AppendChild(ele);
                return new TextRange(_doc, run, ele);
            }
        }
        #endregion
    }
}
