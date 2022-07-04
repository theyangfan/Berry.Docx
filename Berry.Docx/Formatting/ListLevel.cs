using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    public class ListLevel
    {
        private readonly Document _doc;
        private readonly W.Level _level;
        internal ListLevel(Document doc, W.AbstractNum ownerNum, W.Level level)
        {
            _doc = doc;
            _level = level;
        }

        #region Public Properties
        /// <summary>
        /// (eg：1, 1.2, 1.2.3.4.5.6.7.8.9)
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

        public CharacterFormat CharacterFormat => new CharacterFormat(_doc, _level);

        internal float FirstLineIndent
        {
            get
            {
                if(_level.PreviousParagraphProperties?.Indentation == null) return 0;
                W.Indentation ind = _level.PreviousParagraphProperties.Indentation;
                if(ind.Hanging != null) return -(ind.Hanging.Value.ToFloat() / 20);
                if(ind.FirstLine != null) return ind.FirstLine.Value.ToFloat() / 20;
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
        #endregion




        internal int ListLevelNumber => _level.LevelIndex + 1;

        internal string ParagraphStyleId
        {
            get
            {
                if(_level.ParagraphStyleIdInLevel == null) return null;
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
    }
}
