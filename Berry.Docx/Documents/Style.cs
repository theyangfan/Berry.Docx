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

        public CharacterFormat CharacterFormat => _cFormat;

        /// <summary>
        /// 
        /// </summary>
        public StyleType Type
        {
            get => (StyleType)(int)_style.Type.Value;
            set => _style.Type.Value = (W.StyleValues)(int)value;
        }
        /// <summary>
        /// 
        /// </summary>
        public string StyleId
        {
            get => _style.StyleId;
            set => _style.StyleId = value;
        }
        /// <summary>
        /// 
        /// </summary>
        public string Name
        {
            get => _style.StyleName.Val;
            set => _style.StyleName.Val = value;
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
                _style.BasedOn = new W.BasedOn() { Val = value.StyleId};
            }
        }

        /// <summary>
        /// 是否添加到样式库
        /// (This element specifies whether this style shall be treated as a primary style when this document is loaded by an application).
        /// </summary>
        public bool PrimaryStyle
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
