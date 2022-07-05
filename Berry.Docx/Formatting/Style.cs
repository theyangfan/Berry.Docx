// Copyright (c) theyangfan. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

/* Style.cs 文件定义了 Style 类。该类是一个抽象类，是表格、段落、
 * 编号、字符样式的基类。每种样式都具有 Type, Name, Id 等属性。 
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Formatting;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the base class of <see cref="ParagraphStyle"/>、<see cref="CharacterStyle"/>.
    /// <para><see cref="ParagraphStyle"/>、<see cref="CharacterStyle"/> 的基类.</para>
    /// </summary>
    public abstract class Style
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Style _style;
        #endregion

        #region Protected Members
        /// <summary>
        /// The ParagraphFormat of the style.
        /// </summary>
        protected ParagraphFormat _pFormat;
        /// <summary>
        /// The CharacterFormat of the style.
        /// </summary>
        protected CharacterFormat _cFormat;
        /// <summary>
        /// The ListFormat of the style.
        /// </summary>
        protected ListFormat _listFormat;
        #endregion

        #region Constructors
        internal Style(Document doc, StyleType type)
        {
            _doc = doc;
            _style = new W.Style();
            StyleId = IDGenerator.GenerateStyleID(doc);
            Type = type;
            if (type == StyleType.Paragraph)
            {
                _pFormat = new ParagraphFormat(doc, _style);
                _cFormat = new CharacterFormat(doc, _style);
                _listFormat = new ListFormat(doc, _style);
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
                _listFormat = new ListFormat(doc, style);
            }
            else if (Type == StyleType.Character)
            {
                _cFormat = new CharacterFormat(doc, style);
            }
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the CharacterFormat of the style.
        /// </summary>
        public CharacterFormat CharacterFormat => _cFormat;

        /// <summary>
        /// Gets the type of the style.
        /// </summary>
        public StyleType Type
        {
            get => (StyleType)(int)_style.Type.Value;
            private set => _style.Type = (W.StyleValues)(int)value;
        }

        /// <summary>
        /// Gets the unique id of the style.
        /// </summary>
        public string StyleId
        {
            get => _style.StyleId;
            private set => _style.StyleId = value;
        }

        /// <summary>
        /// If current style is the default style, return true; otherwise, return false.
        /// </summary>
        public bool IsDefault => _style.Default ?? false;

        /// <summary>
        /// Gets the name of the current style.
        /// </summary>
        public string Name
        {
            get => _style.StyleName?.Val ?? string.Empty;
            internal set => _style.StyleName = new W.StyleName() { Val = value };
        }

        /// <summary>
        /// Gets or sets the base style of the current style.
        /// </summary>
        public Style BaseStyle
        {
            get
            {
                if (_style.BasedOn != null)
                {
                    return _doc.Styles.Where(s => s.StyleId == _style.BasedOn.Val).FirstOrDefault();
                }
                return null;
            }
            set
            {
                if (value != null)
                    _style.BasedOn = new W.BasedOn() { Val = value.StyleId };
            }
        }

        /// <summary>
        /// If current style is a custom style, return true; otherwise, return false.
        /// </summary>
        public bool IsCustom
        {
            get => _style.CustomStyle ?? false;
            internal set => _style.CustomStyle = value;
        }
        #endregion

        #region Internal Properties
        /// <summary>
        /// This element specifies whether this style shall be treated as a primary style when this document is loaded by an application
        /// <para>是否添加到样式库</para>
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

        /// <summary>
        /// Gets or sets the linked character or paragraph style of the current style.
        /// </summary>
        internal Style LinkedStyle
        {
            get
            {
                if (_style.LinkedStyle != null)
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

        internal W.Style XElement => _style;
        #endregion

        #region Public Methods
        /// <summary>
        /// Converts the string style name to the <see cref="BuiltInStyle"/> type.
        /// </summary>
        /// <param name="styleName">The string style name.</param>
        /// <returns>The <see cref="BuiltInStyle"/> type.</returns>
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
        #endregion

        #region Internal Methods
        /// <summary>
        /// Converts the literal string style name to the built-in string style name.
        /// </summary>
        /// <param name="styleName"></param>
        /// <returns></returns>
        internal static string NameToBuiltInString(string styleName)
        {
            return BuiltInStyleNameFormatter.NameToBuiltInString(styleName);
        }
        #endregion


    }
}
