using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Berry.Docx.VisualModel.Field;

namespace Berry.Docx.VisualModel.Documents
{
    public class ParagraphLine
    {
        private readonly Berry.Docx.Documents.Paragraph _paragraph;
        private double _availableWidth = 0;
        private double _lineSpace = 0;
        private DocGridType _gridType;

        private readonly List<Character> _chars;
        private double _height = 0;
        private double _specialIndent = 0;
        private readonly HorizontalAlignment _hAlign = HorizontalAlignment.Left;

        private double _curWidth = 0;
        private int _rowCnt = 1;

        public ParagraphLine(Berry.Docx.Documents.Paragraph paragraph, double availableWidth, double charSpace, double lineSpace, DocGridType gridType)
        {
            _paragraph = paragraph;
            _availableWidth = availableWidth;
            _lineSpace = lineSpace;
            _gridType = gridType;

            _chars = new List<Character>();
            

            if (gridType != DocGridType.SnapToChars)
            {
                switch (paragraph.Format.Justification)
                {
                    case JustificationType.Left:
                        _hAlign = HorizontalAlignment.Left;
                        break;
                    case JustificationType.Center:
                        _hAlign = HorizontalAlignment.Center;
                        break;
                    case JustificationType.Right:
                        _hAlign = HorizontalAlignment.Right;
                        break;
                    case JustificationType.Both:
                        _hAlign = HorizontalAlignment.Stretch;
                        break;
                    default:
                        break;
                }
            }
        }

        public List<Character> Characters => _chars;
        public double Height => _height;
        public double SpecialIndent
        {
            get => _specialIndent;
            set => _specialIndent = value;
        }

        internal bool HasPageBreak { get; set; }

        public bool TryAppend(Character character)
        {
            if(_specialIndent + _curWidth + character.Width > _availableWidth + 1)
            {
                return false;
            }
            _chars.Add(character);
            _curWidth += character.Width;
            if (_paragraph.Format.SnapToGrid && _gridType != DocGridType.None)
            {
                while (character.Height > (_lineSpace * _rowCnt) * 0.76) _rowCnt++;
                _height = _lineSpace * _rowCnt;
            }
            else
            {
                _height = Math.Max(_height, character.Height);
            }

            var lineSpacing = _paragraph.Format.GetLineSpacing();
            if (_paragraph.Format.SnapToGrid && _gridType != DocGridType.None)
            {
                if (lineSpacing.Rule == LineSpacingRule.Multiple)
                    _height = Math.Max(lineSpacing.Val, _height / _lineSpace) * _lineSpace;
                else if (lineSpacing.Rule == LineSpacingRule.Exactly)
                    _height = lineSpacing.Val.ToPixel();
                else
                    _height = Math.Max(_height, lineSpacing.Val.ToPixel());
            }
            else
            {
                if (lineSpacing.Rule == LineSpacingRule.Multiple)
                    _height = _height * lineSpacing.Val;
                else if (lineSpacing.Rule == LineSpacingRule.Exactly)
                    _height = lineSpacing.Val.ToPixel();
                else
                    _height = Math.Max(_height, lineSpacing.Val.ToPixel());
            }
            return true;
        }
    }
}
