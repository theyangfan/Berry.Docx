// Copyright (c) theyangfan. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the list style level.
    /// <para>表示多级列表样式级别.</para>
    /// </summary>
    public class ListLevel
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Level _level;
        #endregion

        #region Constructors
        internal ListLevel(Document doc, W.AbstractNum ownerNum, W.Level level)
        {
            _doc = doc;
            _level = level;
        }
        #endregion


        #region Public Properties
        /// <summary>
        /// Gets or sets the numbering pattern. The number that higher than the level number will be ignored.
        /// (e.g.: 1, 1.2, 1.2.3.4.5.6.7.8.9)
        /// <para>获取或设置编号的格式。高于当前级别的数字会被忽略.</para>
        /// </summary>
        public string Pattern
        {
            get
            {
                if (_level.LevelText?.Val == null) return string.Empty;
                return _level.LevelText.Val.Value.Replace("%", "");
            }
            set
            {
                if (_level.LevelText == null) _level.LevelText = new W.LevelText();
                _level.LevelText.Val = value.Replace("1", "%1")
                    .Replace("2", "%2")
                    .Replace("3", "%3")
                    .Replace("4", "%4")
                    .Replace("5", "%5")
                    .Replace("6", "%6")
                    .Replace("7", "%7")
                    .Replace("8", "%8")
                    .Replace("9", "%9");
            }
        }

        /// <summary>
        /// Gets or sets the numbering style.
        /// <para>获取或设置编号样式.</para>
        /// </summary>
        public ListNumberStyle NumberStyle
        {
            get
            {
                if (_level.NumberingFormat?.Val == null) return ListNumberStyle.Decimal;
                return _level.NumberingFormat.Val.Value.Convert<ListNumberStyle>();
            }
            set
            {
                if (_level.NumberingFormat == null) _level.NumberingFormat = new W.NumberingFormat();
                _level.NumberingFormat.Val = value.Convert<W.NumberFormatValues>();
            }
        }

        /// <summary>
        /// Gets or sets the start number.
        /// <para>获取或设置起始编号.</para>
        /// </summary>
        public int StartNumber
        {
            get
            {
                if (_level.StartNumberingValue?.Val == null) return 0;
                return _level.StartNumberingValue.Val.Value;
            }
            set
            {
                if (_level.StartNumberingValue == null) _level.StartNumberingValue = new W.StartNumberingValue();
                _level.StartNumberingValue.Val = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether display te current level using arabic numerals.
        /// <para>是否按正规形式(阿拉伯数字)编号.</para>
        /// </summary>
        public bool IsLegalNumberingStyle
        {
            get
            {
                if (_level.IsLegalNumberingStyle == null) return false;
                if (_level.IsLegalNumberingStyle.Val == null) return true;
                return _level.IsLegalNumberingStyle.Val;
            }
            set
            {
                if (value) _level.IsLegalNumberingStyle = new W.IsLegalNumberingStyle();
                else _level.IsLegalNumberingStyle = null;
            }
        }

        /// <summary>
        /// Gets or sets the number alignment.
        /// <para>获取或设置编号对齐方式.</para>
        /// </summary>
        public ListNumberAlignment NumberAlignment
        {
            get
            {
                if(_level.LevelJustification?.Val == null) return ListNumberAlignment.Left;
                return _level.LevelJustification.Val.Value.Convert<ListNumberAlignment>();
            }
            set
            {
                _level.LevelJustification = new W.LevelJustification()
                { 
                    Val = value.Convert<W.LevelJustificationValues>()
                };
            }
        }

        /// <summary>
        /// Gets or sets the character following number.
        /// <para>获取或设置编号之后的符号.</para>
        /// </summary>
        public LevelSuffixCharacter SuffixCharacter
        {
            get
            {
                if (_level.LevelSuffix?.Val == null) return LevelSuffixCharacter.Tab;
                return _level.LevelSuffix.Val.Value.Convert<LevelSuffixCharacter>();
            }
            set
            {
                _level.LevelSuffix = new W.LevelSuffix()
                {
                    Val = value.Convert<W.LevelSuffixValues>()
                };
            }
        }

        /// <summary>
        /// Gets or sets the number align position (in points).
        /// <para>获取或设置编号对齐位置(磅).</para>
        /// </summary>
        public float NumberPosition
        {
            get
            {
                return FirstLineIndent + LeftIndent;
            }
            set
            {
                FirstLineIndent = value - LeftIndent;
            }
        }

        /// <summary>
        /// Gets or sets the text indent position (in points).
        /// <para>获取或设置文本缩进位置(磅).</para>
        /// </summary>
        public float TextIndentation
        {
            get => LeftIndent;
            set
            {
                float temp = NumberPosition;
                LeftIndent = value;
                NumberPosition = temp;
            }
        }

        /// <summary>
        /// Gets the character format.
        /// </summary>
        public CharacterFormat CharacterFormat => new CharacterFormat(_doc, _level);
        #endregion

        #region Internal Properties
        internal float FirstLineIndent
        {
            get
            {
                if (_level.PreviousParagraphProperties?.Indentation == null) return 0;
                W.Indentation ind = _level.PreviousParagraphProperties.Indentation;
                if (ind.Hanging != null) return -(ind.Hanging.Value.ToFloat() / 20);
                if (ind.FirstLine != null) return ind.FirstLine.Value.ToFloat() / 20;
                return 0;
            }
            set
            {
                if (_level.PreviousParagraphProperties == null)
                    _level.PreviousParagraphProperties = new W.PreviousParagraphProperties();
                if (_level.PreviousParagraphProperties.Indentation == null)
                    _level.PreviousParagraphProperties.Indentation = new W.Indentation();
                W.Indentation ind = _level.PreviousParagraphProperties.Indentation;
                if (value >= 0)
                {
                    ind.Hanging = null;
                    ind.FirstLine = ((int)(value * 20)).ToString();
                }
                else
                {
                    ind.FirstLine = null;
                    ind.Hanging = ((int)(-value * 20)).ToString();
                }
            }
        }

        internal float LeftIndent
        {
            get
            {
                if (_level.PreviousParagraphProperties?.Indentation == null) return 0;
                W.Indentation ind = _level.PreviousParagraphProperties.Indentation;
                if (ind.Left != null) return ind.Left.Value.ToFloat() / 20;
                return 0;
            }
            set
            {
                if (_level.PreviousParagraphProperties == null)
                    _level.PreviousParagraphProperties = new W.PreviousParagraphProperties();
                if (_level.PreviousParagraphProperties.Indentation == null)
                    _level.PreviousParagraphProperties.Indentation = new W.Indentation();
                W.Indentation ind = _level.PreviousParagraphProperties.Indentation;
                ind.Left = ((int)(value * 20)).ToString();
            }
        }

        internal int ListLevelNumber => _level.LevelIndex + 1;

        internal string ParagraphStyleId
        {
            get
            {
                if (_level.ParagraphStyleIdInLevel == null) return null;
                return _level.ParagraphStyleIdInLevel.Val;
            }
            set
            {
                if (value == null)
                    _level.ParagraphStyleIdInLevel = null;
                else
                    _level.ParagraphStyleIdInLevel = new W.ParagraphStyleIdInLevel() { Val = value };
            }
        }
        #endregion
    }
}
