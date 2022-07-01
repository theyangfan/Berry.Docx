using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Formatting;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// TODO
    /// </summary>
    public abstract class Style
    {
        private readonly Document _doc;
        private readonly W.Style _style;
        protected ParagraphFormat _pFormat;
        protected CharacterFormat _cFormat;

        internal Style(Document doc, StyleType type)
        {
            _doc = doc;
            _style = new W.Style();
            StyleId = IDGenerator.GenerateStyleID(doc);
            Type = type;
            if (type == StyleType.Paragraph || type == StyleType.Table)
            {
                _pFormat = new ParagraphFormat(doc, _style);
                _cFormat = new CharacterFormat(doc, _style);
            }
            else if (type == StyleType.Character)
            {
                _cFormat = new CharacterFormat(doc, _style);
            }
        }

        internal Style(Document doc, W.Style style)
        {
            _doc = doc;
            _style = style;
            if(Type == StyleType.Paragraph)
            {
                _pFormat = new ParagraphFormat(doc, style);
                _cFormat = new CharacterFormat(doc, style);
            }
            else if (Type == StyleType.Character)
            {
                _cFormat = new CharacterFormat(doc, style);
            }
        }

        internal W.Style XElement => _style;
        public CharacterFormat CharacterFormat => _cFormat;

        /// <summary>
        /// 
        /// </summary>
        public StyleType Type
        {
            get => (StyleType)(int)_style.Type.Value;
            private set => _style.Type = (W.StyleValues)(int)value;
        }
        /// <summary>
        /// 
        /// </summary>
        public string StyleId
        {
            get => _style.StyleId;
            private set => _style.StyleId = value;
        }

        public bool IsDefault => _style.Default ?? false;

        /// <summary>
        /// 
        /// </summary>
        public string Name
        {
            get => _style.StyleName?.Val ?? string.Empty;
            internal set => _style.StyleName = new W.StyleName() { Val = value };
        }

        public Style BaseStyle
        {
            get
            {
                if(_style.BasedOn != null)
                {
                    return _doc.Styles.Where(s => s.StyleId == _style.BasedOn.Val).FirstOrDefault();
                }
                return null;
            }
            set
            {
                if(value != null)
                    _style.BasedOn = new W.BasedOn() { Val = value.StyleId};
            }
        }

        public bool IsCustom
        {
            get => _style.CustomStyle ?? false;
            internal set => _style.CustomStyle = value;
        }

        /// <summary>
        /// 是否添加到样式库
        /// (This element specifies whether this style shall be treated as a primary style when this document is loaded by an application).
        /// </summary>
        internal bool AddToGallery
        {
            get
            {
                if (_style.PrimaryStyle == null) return false;
                if (_style.PrimaryStyle.Val == null) return true;
                return _style.PrimaryStyle.Val.Value == W.OnOffOnlyValues.On;
            }
            set
            {
                if (value)
                {
                    if (_style.PrimaryStyle == null)
                        _style.PrimaryStyle = new W.PrimaryStyle();
                    else
                        _style.PrimaryStyle.Val = null;
                }
                else
                {
                    _style.PrimaryStyle = null;
                }
            }
        }

        internal Style LinkedStyle
        {
            get
            {
                if(_style.LinkedStyle != null)
                {
                    string id = _style.LinkedStyle.Val;
                    return _doc.Styles.Where(s => s.StyleId == id).FirstOrDefault();
                }
                return null;
            }
            set
            {
                _style.LinkedStyle = new W.LinkedStyle() { Val = value.StyleId };
            }
        }

        public static BuiltInStyle NameToBuiltIn(string styleName)
        {
            styleName = NameToBuiltInString(styleName);
            if (styleName == "normal") return BuiltInStyle.Normal;
            else if (styleName == "heading 1") return BuiltInStyle.Heading1;
            else if (styleName == "heading 2") return BuiltInStyle.Heading2;
            else if (styleName == "heading 3") return BuiltInStyle.Heading3;
            else if (styleName == "heading 4") return BuiltInStyle.Heading4;
            else if (styleName == "heading 5") return BuiltInStyle.Heading5;
            else if (styleName == "heading 6") return BuiltInStyle.Heading6;
            else if (styleName == "heading 7") return BuiltInStyle.Heading7;
            else if (styleName == "heading 8") return BuiltInStyle.Heading8;
            else if (styleName == "heading 9") return BuiltInStyle.Heading9;
            else if (styleName == "toc 1") return BuiltInStyle.TOC1;
            else if (styleName == "toc 2") return BuiltInStyle.TOC2;
            else if (styleName == "toc 3") return BuiltInStyle.TOC3;
            else return BuiltInStyle.None;
        }
        internal static string NameToBuiltInString(string styleName)
        {
            styleName = styleName.ToLower();
            if (styleName == "正文") return "normal";
            else if (styleName == "标题 1") return "heading 1";
            else if (styleName == "标题 2") return "heading 2";
            else if (styleName == "标题 3") return "heading 3";
            else if (styleName == "标题 4") return "heading 4";
            else if (styleName == "标题 5") return "heading 5";
            else if (styleName == "标题 6") return "heading 6";
            else if (styleName == "标题 7") return "heading 7";
            else if (styleName == "标题 8") return "heading 8";
            else if (styleName == "标题 9") return "heading 9";
            else if (styleName == "目录 1") return "toc 1";
            else if (styleName == "目录 2") return "toc 2";
            else if (styleName == "目录 3") return "toc 3";
            else return styleName;
        }

    }
}
