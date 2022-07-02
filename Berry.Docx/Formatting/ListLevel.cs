using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    public class ListLevel
    {
        private readonly W.Level _level;
        internal ListLevel(Document doc, W.AbstractNum ownerNum, W.Level level)
        {
            _level = level;
        }
        /// <summary>
        /// (eg：1, 1.2, 1.2.3.4.5.6.7.8.9)
        /// </summary>
        public string Pattern
        {
            get
            {
                if(_level.LevelText?.Val == null) return string.Empty;
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
                if(_level.NumberingFormat == null) _level.NumberingFormat = new W.NumberingFormat();
                _level.NumberingFormat.Val = value.Convert<W.NumberFormatValues>();
            }
        }


        public int StartNumber
        {
            get
            {
                if(_level.StartNumberingValue?.Val == null) return 0;
                return _level.StartNumberingValue.Val.Value;
            }
            set
            {
                if (_level.StartNumberingValue == null) _level.StartNumberingValue = new W.StartNumberingValue();
                _level.StartNumberingValue.Val = value;
            }
        }

    }
}
