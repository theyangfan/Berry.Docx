using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Formatting;

namespace Berry.Docx.Documents
{
    public class Style
    {
        private W.Style _style = null;
        protected ParagraphFormat _pFormat = null;
        protected CharacterFormat _cFormat = null;
        public Style(Document doc, W.Style style)
        {
            _style = style;
            if(style.Type.Value == W.StyleValues.Paragraph)
            {
                _pFormat = new ParagraphFormat(doc, style);
                _cFormat = new CharacterFormat(doc, style);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public StyleType Type
        {
            get
            {
                if (_style.Type.Value == W.StyleValues.Paragraph)
                    return StyleType.Paragraph;
                else if (_style.Type.Value == W.StyleValues.Character)
                    return StyleType.Character;
                else if (_style.Type.Value == W.StyleValues.Table)
                    return StyleType.Table;
                else 
                    return StyleType.Numbering;
            }
            set
            {
                if (value == StyleType.Paragraph)
                    _style.Type.Value = W.StyleValues.Paragraph;
                else if (value == StyleType.Character)
                    _style.Type.Value = W.StyleValues.Character;
                else if (value == StyleType.Table)
                    _style.Type.Value = W.StyleValues.Table;
                else
                    _style.Type.Value = W.StyleValues.Numbering;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public string StyleId
        {
            get
            {
                return _style.StyleId;
            }
            set
            {
                _style.StyleId = value;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public string Name
        {
            get
            {
                return _style.StyleName.Val;
            }
            set
            {
                _style.StyleName.Val = value;
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
