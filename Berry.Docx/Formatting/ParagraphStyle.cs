﻿// Copyright (c) theyangfan. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Documents;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the paragraph style.
    /// <para>表示一个段落样式，支持读写其字符和段落格式。可以通过 <c>CreateBuiltInStyle</c> 静态方法创建指定的内置样式。</para>
    /// </summary>
    public class ParagraphStyle : Style
    {
        #region Private Members
        private readonly Document _doc;
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a paragraph style with the specified name which is based on normal style.
        /// <para>创建一个指定名称的段落样式，其基于正文样式。</para>
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="styleName">样式名</param>
        public ParagraphStyle(Document doc, string styleName) : this(doc, styleName, BuiltInStyle.Normal){}

        /// <summary>
        /// Creates a paragraph style with the specified name which is based on the specified style.
        /// <para>创建一个基于指定样式的指定名称的段落样式。</para>
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="styleName">样式名</param>
        /// <param name="basedStyle">基类样式</param>
        public ParagraphStyle(Document doc, string styleName, BuiltInStyle basedStyle) : base(doc, StyleType.Paragraph)
        {
            _doc = doc;
            this.Name = styleName;
            this.IsCustom = true;
            this.AddToGallery = true;
            if(basedStyle != BuiltInStyle.None)
            {
                this.BaseStyle = CreateBuiltInStyle(basedStyle, doc);
            }
        }

        internal ParagraphStyle(Document doc, W.Style style):base(doc, style)
        {
            _doc = doc;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the paragraph format.
        /// <para>访问当前样式的段落格式.</para>
        /// </summary>
        public ParagraphFormat ParagraphFormat => _pFormat;

        /// <summary>
        /// Gets the list format.
        /// <para>访问当前样式的列表格式.</para>
        /// </summary>
        public ListFormat ListFormat => _listFormat;

        /// <summary>
        /// Gets or sets the base style.
        /// <para>获取或设置基样式.</para>
        /// </summary>
        public new ParagraphStyle BaseStyle
        {
            get => base.BaseStyle as ParagraphStyle;
            set => base.BaseStyle = value;
        }
        #endregion

        public static ParagraphStyle Default(Document doc)
        {
            return doc.Styles.Where(s => s.Type == StyleType.Paragraph && s.IsDefault).FirstOrDefault() as ParagraphStyle;
        }

        public CharacterStyle GetLinkedStyle()
        {
            return base.LinkedStyle as CharacterStyle;
        }

        public CharacterStyle CreateLinkedStyle()
        {
            if(base.LinkedStyle != null)
            {
                return base.LinkedStyle as CharacterStyle;
            }
            CharacterStyle linked = new CharacterStyle(_doc);
            linked.Name = this.Name.Replace("heading", "标题") + " 字符";
            linked.BaseStyle = CharacterStyle.Default(_doc);
            linked.LinkedStyle = this;
            linked.IsCustom = true;
            this.LinkedStyle = linked;
            // copy character format
            linked.CharacterFormat.FontNameEastAsia = this.CharacterFormat.FontNameEastAsia;
            linked.CharacterFormat.FontNameAscii = this.CharacterFormat.FontNameAscii;
            linked.CharacterFormat.FontSize = this.CharacterFormat.FontSize;
            linked.CharacterFormat.FontSizeCs = this.CharacterFormat.FontSizeCs;
            linked.CharacterFormat.Bold = this.CharacterFormat.Bold;
            linked.CharacterFormat.Italic = this.CharacterFormat.Italic;
            linked.CharacterFormat.CharacterScale = this.CharacterFormat.CharacterScale;
            linked.CharacterFormat.CharacterSpacing = this.CharacterFormat.CharacterSpacing;
            linked.CharacterFormat.Position = this.CharacterFormat.Position;
            // add to style list
            _doc.Styles.Add(linked);
            return linked;
        }

        public static ParagraphStyle CreateBuiltInStyle(BuiltInStyle bstyle, Document doc)
        {
            string styleName = string.Empty;
            switch (bstyle)
            {
                case BuiltInStyle.Normal:
                    styleName = "Normal";
                    break;
                case BuiltInStyle.Heading1:
                    styleName = "heading 1";
                    break;
                case BuiltInStyle.Heading2:
                    styleName = "heading 2";
                    break;
                case BuiltInStyle.Heading3:
                    styleName = "heading 3";
                    break;
                case BuiltInStyle.Heading4:
                    styleName = "heading 4";
                    break;
                case BuiltInStyle.Heading5:
                    styleName = "heading 5";
                    break;
                case BuiltInStyle.Heading6:
                    styleName = "heading 6";
                    break;
                case BuiltInStyle.Heading7:
                    styleName = "heading 7";
                    break;
                case BuiltInStyle.Heading8:
                    styleName = "heading 8";
                    break;
                case BuiltInStyle.Heading9:
                    styleName = "heading 9";
                    break;
                default:
                    break;
            }
            if (string.IsNullOrEmpty(styleName)) return null;

            if (doc.Styles.FindByName(styleName, StyleType.Paragraph) != null)
            {
                return doc.Styles.FindByName(styleName, StyleType.Paragraph) as ParagraphStyle;
            }
            else
            {
                W.Style s = BuiltInStyleGenerator.Generate(doc, bstyle);
                if(s == null) return null;
                ParagraphStyle style = new ParagraphStyle(doc, s);
                doc.Styles.Add(style);
                return style;
            }
        }

        


    }
}
