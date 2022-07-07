// Copyright (c) theyangfan. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the character style. Normally, the character style links to a paragraph style, rather than comes alone.
    /// <para>表示一个字符样式。一般情况下，字符样式会链接到段落样式，而非单独出现。</para>
    /// </summary>
    public class CharacterStyle : Style
    {
        #region Constructors
        internal CharacterStyle(Document doc) : base(doc, StyleType.Character)
        {
        }
        internal CharacterStyle(Document doc, W.Style style) : base(doc, style)
        {
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets or sets the base style.
        /// <para>获取或设置基样式.</para>
        /// </summary>
        public new CharacterStyle BaseStyle
        {
            get => base.BaseStyle as CharacterStyle;
            set => base.BaseStyle = value;
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Gets the default character style.
        /// <para>获取默认字符样式.</para>
        /// </summary>
        /// <param name="doc"></param>
        /// <returns>The default paragraph style.</returns>
        public static CharacterStyle Default(Document doc)
        {
            return doc.Styles.Where(s => s.Type == StyleType.Character && s.IsDefault).FirstOrDefault() as CharacterStyle;
        }
        #endregion

    }
}
