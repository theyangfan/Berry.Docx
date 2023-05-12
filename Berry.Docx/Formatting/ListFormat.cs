// Copyright (c) theyangfan. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the numbering format. You can access or specify the <see cref="ListStyle"/> of the current paragraph/style.
    /// <para>表示段落或段落样式的编号格式. 可以访问或指定当前段落/样式的多级列表样式.</para>
    /// </summary>
    public class ListFormat
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Paragraph _ownerParagraph;
        private readonly W.Style _ownerStyle;
        #endregion

        #region Constructors
        internal ListFormat(Document doc, W.Paragraph ownerParagraph)
        {
            _doc = doc;
            _ownerParagraph = ownerParagraph;
        }

        internal ListFormat(Document doc, W.Style ownerStyle)
        {
            _doc = doc;
            _ownerStyle = ownerStyle;
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// Gets the current list style(may not exist).
        /// <para>获取当前列表样式(可能不存在).</para>
        /// </summary>
        public ListStyle CurrentStyle
        {
            get
            {
                if(_ownerParagraph != null) // Paragraph
                {
                    // direct formatting
                    W.NumberingProperties numPr = _ownerParagraph.ParagraphProperties?.NumberingProperties;
                    if (numPr?.NumberingId != null && numPr?.NumberingLevelReference != null)
                    {
                        W.AbstractNum abstractNum = GetAbstractNumByID(_doc, numPr.NumberingId.Val);
                        if(abstractNum != null) return new ListStyle(_doc, abstractNum);
                    }
                    else // inherit from paragraph style
                    {
                        W.AbstractNum abstractNum = GetStyleAbstractNumRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                        if (abstractNum != null) return new ListStyle(_doc, abstractNum);
                    }
                }
                else // Style
                {
                    W.AbstractNum abstractNum = GetStyleAbstractNumRecursively(_doc, _ownerStyle);
                    if (abstractNum != null) return new ListStyle(_doc, abstractNum);
                }
                return null;
            }
        }

        /// <summary>
        /// Gets the current list level format(may not exist).
        /// <para>获取当前列表级别格式(可能不存在).</para>
        /// </summary>
        public ListLevel CurrentLevel
        {
            get
            {
                if (_ownerParagraph != null) // Paragraph
                {
                    // direct formatting
                    W.NumberingProperties numPr = _ownerParagraph.ParagraphProperties?.NumberingProperties;
                    if (numPr?.NumberingId != null && numPr?.NumberingLevelReference != null)
                    {
                        W.AbstractNum abstractNum = GetAbstractNumByID(_doc, numPr.NumberingId.Val);
                        if (abstractNum != null)
                        {
                            W.Level level = GetAbstractNumLevel(abstractNum, numPr.NumberingLevelReference.Val.Value);
                            if(level != null) return new ListLevel(_doc, abstractNum, level);
                        }
                    }
                    else // inherit from paragraph style
                    {
                        W.Style style = _ownerParagraph.GetStyle(_doc);
                        W.AbstractNum abstractNum = GetStyleAbstractNumRecursively(_doc, style);
                        if (abstractNum != null)
                        {
                            W.Level level = GetAbstractNumLevel(abstractNum, style.StyleId);
                            if (level != null) return new ListLevel(_doc, abstractNum, level);
                        }
                    }
                }
                else // Style
                {
                    W.AbstractNum abstractNum = GetStyleAbstractNumRecursively(_doc, _ownerStyle);
                    if (abstractNum != null)
                    {
                        W.Level level = GetAbstractNumLevel(abstractNum, _ownerStyle.StyleId);
                        if (level != null) return new ListLevel(_doc, abstractNum, level);
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Gets or sets the current list level number (from 1 to 9). Return zero if not exist.
        /// <para>获取或设置当前列表级别编号(从1-9)。如果不存在，则返回0.</para>
        /// </summary>
        public int ListLevelNumber
        {
            get
            {
                return CurrentLevel?.ListLevelNumber ?? 0;
            }
            set
            {
                InitNumberingProperties();
                if (_ownerParagraph != null)
                {
                    // set the numbering level index to the direct formatting
                    W.NumberingProperties num = _ownerParagraph.ParagraphProperties.NumberingProperties;
                    num.NumberingLevelReference = new W.NumberingLevelReference() { Val = value - 1 };
                }
                else
                {
                    if (CurrentStyle == null) return;
                    foreach (ListLevel level in CurrentStyle.Levels)
                    {
                        if (level.ParagraphStyleId == _ownerStyle.StyleId)
                            level.ParagraphStyleId = null;
                    }
                    // set the style id to the numbering level
                    CurrentStyle.Levels[value - 1].ParagraphStyleId = _ownerStyle.StyleId;
                }
            }
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Apply the specified list style and level.
        /// <para>应用指定的列表样式和级别.</para>
        /// </summary>
        /// <param name="style">The list style</param>
        /// <param name="levelNumber">The list level number.</param>
        public void ApplyStyle(ListStyle style, int levelNumber)
        {
            InitNumberingProperties();
            if (_ownerParagraph != null)
            {
                W.NumberingProperties num = _ownerParagraph.ParagraphProperties.NumberingProperties;
                num.NumberingId = new W.NumberingId() { Val = style.NumberID };
                num.NumberingLevelReference = new W.NumberingLevelReference() { Val = levelNumber - 1 };
            }
            else
            {
                // Add properties to style
                W.NumberingProperties num = _ownerStyle.StyleParagraphProperties.NumberingProperties;
                num.NumberingId = new W.NumberingId() { Val = style.NumberID };
                num.NumberingLevelReference = new W.NumberingLevelReference() { Val = levelNumber - 1 };
                // Add properties to list
                foreach (ListLevel level in style.Levels)
                {
                    if (level.ParagraphStyleId == _ownerStyle.StyleId)
                        level.ParagraphStyleId = null;
                }
                style.Levels[levelNumber - 1].ParagraphStyleId = _ownerStyle.StyleId;
            }
        }

        /// <summary>
        /// Apply the list style with the specified name. The list style name does not exist physically, the name will be invalid when out of document scope.
        /// <para>应用指定名称的列表样式。列表样式的名称在物理上不存在，当离开文档作用域后，名称将无效。</para>
        /// </summary>
        /// <param name="styleName">The list style name.</param>
        /// <param name="levelNumber">The list level number.</param>
        public void ApplyStyle(string styleName, int levelNumber)
        {
            ListStyle style = _doc.ListStyles.FindByName(styleName);
            if (style == null) return;
            ApplyStyle(style, levelNumber);
        }

        /// <summary>
        /// Clears the list format.
        /// </summary>
        public void ClearFormatting()
        {
            if(_ownerParagraph != null)
            {
                var style = new Paragraph(_doc, _ownerParagraph).GetStyle();
                if(style?.ListFormat.CurrentStyle != null)
                {
                    if (_ownerParagraph.ParagraphProperties == null)
                        _ownerParagraph.ParagraphProperties = new W.ParagraphProperties();
                    _ownerParagraph.ParagraphProperties.NumberingProperties = new W.NumberingProperties()
                    {
                        NumberingId = new W.NumberingId() { Val = 0 }
                    };
                }
                else if(_ownerParagraph.ParagraphProperties?.NumberingProperties != null)
                {
                    _ownerParagraph.ParagraphProperties.NumberingProperties = null;
                }
            }
            else if(_ownerStyle != null)
            {
                var baseStyle = new ParagraphStyle(_doc, _ownerStyle).BaseStyle;
                if(baseStyle?.ListFormat?.CurrentStyle != null)
                {
                    if (_ownerStyle.StyleParagraphProperties == null)
                        _ownerStyle.StyleParagraphProperties = new W.StyleParagraphProperties();
                    _ownerStyle.StyleParagraphProperties.NumberingProperties = new W.NumberingProperties()
                    {
                        NumberingId = new W.NumberingId() { Val = 0 }
                    };
                }
                else if (_ownerStyle.StyleParagraphProperties?.NumberingProperties != null)
                {
                    _ownerStyle.StyleParagraphProperties.NumberingProperties = null;
                }
            }
        }

        #endregion

        #region Private Methods

        private void InitNumberingProperties()
        {
            if (_ownerParagraph != null)
            {
                if (_ownerParagraph.ParagraphProperties == null)
                    _ownerParagraph.ParagraphProperties = new W.ParagraphProperties();
                if (_ownerParagraph.ParagraphProperties.NumberingProperties == null)
                    _ownerParagraph.ParagraphProperties.NumberingProperties = new W.NumberingProperties();
            }
            else
            {
                if (_ownerStyle.StyleParagraphProperties == null)
                    _ownerStyle.StyleParagraphProperties = new W.StyleParagraphProperties();
                if (_ownerStyle.StyleParagraphProperties.NumberingProperties == null)
                    _ownerStyle.StyleParagraphProperties.NumberingProperties = new W.NumberingProperties();
            }
        }
        private static W.AbstractNum GetAbstractNumByID(Document doc, int numId)
        {
            W.Numbering numbering = doc.Package.MainDocumentPart.NumberingDefinitionsPart?.Numbering;
            if (numbering == null) return null;
            W.NumberingInstance num = numbering.Elements<W.NumberingInstance>().Where(n => n.NumberID == numId).FirstOrDefault();
            if (num == null) return null;
            int abstractNumId = num.AbstractNumId.Val;
            W.AbstractNum abstractNum = numbering.Elements<W.AbstractNum>().Where(a => a.AbstractNumberId == abstractNumId).FirstOrDefault();
            return abstractNum;
        }

        private static W.Level GetAbstractNumLevel(W.AbstractNum num, int levelIndex)
        {
            return num?.Elements<W.Level>().Where(l => l.LevelIndex == levelIndex).FirstOrDefault();
        }

        private static W.Level GetAbstractNumLevel(W.AbstractNum num, string styleId)
        {
            return num?.Elements<W.Level>().Where(l => l.ParagraphStyleIdInLevel?.Val == styleId).FirstOrDefault();
        }

        private static W.AbstractNum GetStyleAbstractNumRecursively(Document doc, W.Style style)
        {
            if (style.StyleParagraphProperties?.NumberingProperties?.NumberingId != null)
            {
                int numId = style.StyleParagraphProperties.NumberingProperties.NumberingId.Val;
                W.AbstractNum abstractNum = GetAbstractNumByID(doc, numId);
                if(abstractNum != null) return abstractNum;
            }
            W.Style baseStyle = style.GetBaseStyle(doc);
            if (baseStyle != null)
            {
                return GetStyleAbstractNumRecursively(doc, baseStyle);
            }
            return null;
        }
        #endregion

    }
}
