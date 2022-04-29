using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Formatting;
using Berry.Docx.Utils;

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

        public Style(Document doc, StyleType type)
        {
            _style = new W.Style();
            StyleId = IDGenerator.GenerateStyleID(doc);
            Type = type;
            if (type == StyleType.Paragraph)
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
            set => _style.StyleName = new W.StyleName() { Val = value };
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
            set => _style.CustomStyle = value;
        }

        /// <summary>
        /// 是否添加到样式库
        /// (This element specifies whether this style shall be treated as a primary style when this document is loaded by an application).
        /// </summary>
        public bool AddToGallery
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


    }
}
