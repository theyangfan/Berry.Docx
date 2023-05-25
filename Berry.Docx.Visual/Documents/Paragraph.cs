using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Windows;
using Berry.Docx.Visual.Field;

namespace Berry.Docx.Visual.Documents
{
    public class Paragraph
    {
        private Berry.Docx.Documents.Paragraph _paragraph;
        private readonly double _width = 0;
        private double _charSpace = 0;
        private double _lineSpace = 0;
        private DocGridType _gridType;
        
        private double _leftIndent = 0;
        private double _rightIndent = 0;
        private double _specialIndent = 0;
        private double _beforeSpace = 0;
        private double _afterSpace = 0;
        private Margin _margin;

        private double _normalFontSize = 10.5;

        private List<ParagraphLine> _lines;

        internal Paragraph(Berry.Docx.Documents.Paragraph paragraph, double availableWidth, double charSpace, double lineSpace, Berry.Docx.DocGridType gridType)
        {
            _paragraph = paragraph;
            _width = availableWidth;
            _charSpace = charSpace;
            _lineSpace = lineSpace;
            _gridType = gridType;

            _lines = new List<ParagraphLine>();

            float firstCharSize = 0;
            if (paragraph.ChildItems.OfType<Berry.Docx.Field.TextRange>().Count() > 0)
            {
                firstCharSize = paragraph.ChildItems.OfType<Berry.Docx.Field.TextRange>().First().CharacterFormat.FontSize;
            }
            else
            {
                firstCharSize = paragraph.MarkFormat.FontSize;
                paragraph.ChildItems.InsertAt(new Berry.Docx.Field.TextRange(_paragraph.Document, " "), 0);
            }

            _normalFontSize = Berry.Docx.Formatting.ParagraphStyle.Default(paragraph.Document).CharacterFormat.FontSize;

            #region 缩进
            // 缩进
            var leftInd = paragraph.Format.GetLeftIndent();
            var rightInd = paragraph.Format.GetRightIndent();
            var specialInd = paragraph.Format.GetSpecialIndentation();
            // 左侧缩进
            if (leftInd.Unit == IndentationUnit.Character)
            {
                if (gridType == DocGridType.LinesAndChars)
                {
                    _leftIndent = charSpace * leftInd.Val;
                }
                else if (gridType == DocGridType.SnapToChars)
                {
                    _leftIndent = charSpace * Math.Ceiling(leftInd.Val);
                }
                else
                {
                    _leftIndent = (_normalFontSize * leftInd.Val) / 72 * 96;
                }
            }
            else
            {
                _leftIndent = leftInd.Val / 72 * 96;
                if (gridType == DocGridType.SnapToChars)
                {
                    int cnt = 1;
                    while (_leftIndent > charSpace * cnt) cnt++;
                    _leftIndent = charSpace * cnt;
                }
            }
            // 右侧缩进
            if (rightInd.Unit == IndentationUnit.Character)
            {
                if (gridType == DocGridType.LinesAndChars
                    || gridType == DocGridType.SnapToChars)
                {
                    _rightIndent = charSpace * rightInd.Val;
                }
                else if (gridType == DocGridType.SnapToChars)
                {
                    _rightIndent = charSpace * Math.Ceiling(rightInd.Val);
                }
                else
                {
                    _rightIndent = (_normalFontSize * rightInd.Val) / 72 * 96;
                }
            }
            else
            {
                _rightIndent = rightInd.Val / 72 * 96;
                if (gridType == DocGridType.SnapToChars)
                {
                    int cnt = 1;
                    while (_rightIndent > charSpace * cnt) cnt++;
                    _rightIndent = charSpace * cnt;
                }
            }

            // 特殊缩进
            if (gridType == DocGridType.LinesAndChars)
            {
                if (specialInd.Type == SpecialIndentationType.FirstLine)
                {
                    if (specialInd.Unit == IndentationUnit.Character)
                        _specialIndent = (charSpace + (firstCharSize - _normalFontSize) / 72 * 96) * specialInd.Val;
                    else
                        _specialIndent = specialInd.Val / 72 * 96;
                }
                else if (specialInd.Type == SpecialIndentationType.Hanging)
                {
                    if (specialInd.Unit == IndentationUnit.Character)
                        _specialIndent = -(charSpace + (firstCharSize - _normalFontSize) / 72 * 96) * specialInd.Val;
                    else
                        _specialIndent = -specialInd.Val / 72 * 96;
                }
            }
            else if (gridType == DocGridType.SnapToChars)
            {
                if (specialInd.Type == SpecialIndentationType.FirstLine)
                {
                    if (specialInd.Unit == IndentationUnit.Character)
                        _specialIndent = Math.Ceiling((charSpace + (firstCharSize - _normalFontSize) / 72 * 96) * specialInd.Val / charSpace) * charSpace;
                    else
                        _specialIndent = Math.Ceiling(specialInd.Val / 72 * 96 / charSpace) * charSpace;
                }
                else if (specialInd.Type == SpecialIndentationType.Hanging)
                {
                    if (specialInd.Unit == IndentationUnit.Character)
                        _specialIndent = -Math.Ceiling((charSpace + (firstCharSize - _normalFontSize) / 72 * 96) * specialInd.Val / charSpace) * charSpace;
                    else
                        _specialIndent = -Math.Ceiling(specialInd.Val / 72 * 96 / charSpace) * charSpace;
                }
            }
            else
            {
                if (specialInd.Type == SpecialIndentationType.FirstLine)
                {
                    if (specialInd.Unit == IndentationUnit.Character)
                        _specialIndent = firstCharSize / 72 * 96 * specialInd.Val;
                    else
                        _specialIndent = specialInd.Val / 72 * 96;
                }
                else if (specialInd.Type == SpecialIndentationType.Hanging)
                {
                    if (specialInd.Unit == IndentationUnit.Character)
                        _specialIndent = -firstCharSize / 72 * 96 * specialInd.Val;
                    else
                        _specialIndent = -specialInd.Val / 72 * 96;
                }
            }
            #endregion

            #region 间距
            // 段前段后间距
            var beforeSpacing = paragraph.Format.GetBeforeSpacing();
            var afterSpacing = paragraph.Format.GetAfterSpacing();
            if (gridType == DocGridType.None)
            {
                if (beforeSpacing.Unit == SpacingUnit.Line)
                    _beforeSpace = beforeSpacing.Val * 12f.ToPixel();
                else
                    _beforeSpace = beforeSpacing.Val.ToPixel();

                if (afterSpacing.Unit == SpacingUnit.Line)
                    _afterSpace = afterSpacing.Val * 12f.ToPixel();
                else
                    _afterSpace = afterSpacing.Val.ToPixel();
            }
            else
            {
                if (beforeSpacing.Unit == SpacingUnit.Line)
                    _beforeSpace = beforeSpacing.Val * lineSpace;
                else
                    _beforeSpace = beforeSpacing.Val.ToPixel();

                if (afterSpacing.Unit == SpacingUnit.Line)
                    _afterSpace = afterSpacing.Val * lineSpace;
                else
                    _afterSpace = afterSpacing.Val.ToPixel();
            }
            #endregion

            _margin = new Margin(_leftIndent, _beforeSpace, _rightIndent, _afterSpace);
        }

        public double Width => _width;

        public List<ParagraphLine> Lines => _lines;

        public Margin Margin => _margin;

        internal List<ParagraphLine> GenerateLines()
        {
            List<ParagraphLine> lines = new List<ParagraphLine>();
            int index = 0;
            lines.Add(new ParagraphLine(_paragraph, _width - _leftIndent - _rightIndent, _charSpace, _lineSpace, _gridType));
            if (_specialIndent > 0) lines[index].SpecialIndent = _specialIndent;

            foreach (var item in _paragraph.ChildItems)
            {
                if (item is Berry.Docx.Field.TextRange)
                {
                    var tr = (Berry.Docx.Field.TextRange)item;
                    foreach (var c in tr.Characters)
                    {
                        Character character = new Character(c, _charSpace, _normalFontSize, _gridType);
                        if (!lines[index].TryAppend(character))
                        {
                            var line = new ParagraphLine(_paragraph, _width - _leftIndent - _rightIndent, _charSpace, _lineSpace, _gridType);
                            if (_specialIndent < 0) line.SpecialIndent = Math.Abs(_specialIndent);
                            lines.Add(line);
                            index++;
                        }
                    }
                }
                else if (item is Berry.Docx.Field.Break)
                {
                    var br = (Berry.Docx.Field.Break)item;
                    if (br.Type == BreakType.Page)
                    {
                        lines[index].HasPageBreak = true;
                        var line = new ParagraphLine(_paragraph, _width - _leftIndent - _rightIndent, _charSpace, _lineSpace, _gridType);
                        if (item == _paragraph.ChildItems.First())
                        {
                            if (_specialIndent > 0) line.SpecialIndent = _specialIndent;
                        }
                        else if (item == _paragraph.ChildItems.Last())
                        {
                            continue;
                        }
                        else
                        {
                            if (_specialIndent < 0) line.SpecialIndent = Math.Abs(_specialIndent);
                        }
                        lines.Add(line);
                        index++;
                    }
                }

            }
            return lines;
        }

    }
}
