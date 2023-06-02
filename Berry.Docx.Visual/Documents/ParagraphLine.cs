using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Berry.Docx.Visual.Field;

namespace Berry.Docx.Visual.Documents
{
    public class ParagraphLine
    {
        #region Private Members
        private readonly Berry.Docx.Documents.Paragraph _paragraph;
        private double _availableWidth = 0;
        private double _lineSpace = 0;
        private DocGridType _gridType;

        private readonly List<ParagraphItem> _items;
        private double _height = 0;
        private readonly HorizontalAlignment _hAlign = HorizontalAlignment.Left;

        private double _textWidth = 0;
        private int _rowCnt = 1;
        private double _maxCharHeight = 0;

        private Margin _margin = new Margin(0, 0, 0, 0);
        private Margin _padding = new Margin(0, 0, 0, 0);
        #endregion

        #region Constructor
        internal ParagraphLine(Berry.Docx.Documents.Paragraph paragraph, double availableWidth, double charSpace, double lineSpace, DocGridType gridType)
        {
            _paragraph = paragraph;
            _availableWidth = availableWidth;
            _lineSpace = lineSpace;
            _gridType = gridType;

            _items = new List<ParagraphItem>();
            
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
                        _hAlign = HorizontalAlignment.Left;
                        break;
                    case JustificationType.Distribute:
                        _hAlign = HorizontalAlignment.Stretch;
                        break;
                    default:
                        break;
                }
            }
        }
        #endregion

        #region Public Properties
        public double Width => _availableWidth;

        public double Height => _height;

        public Margin Margin => _margin;

        public Margin Padding => _padding;

        public HorizontalAlignment HorizontalAlignment => _hAlign;

        public List<ParagraphItem> ChildItems => _items;
        #endregion

        #region Internal Properties
        internal bool EndsWithPageBreak { get; set; }
        #endregion

        #region Internal Methods
        internal bool TryAppend(ParagraphItem item)
        {
            double space = _margin.Left + _margin.Right + _padding.Left + _padding.Right;
            if(item.GetType() != typeof(Picture) || _items.Count > 0)
            {
                if (space + _textWidth + item.Width > _availableWidth + 1)
                {
                    return false;
                }
            }
            _items.Add(item);
            _textWidth += item.Width;
            if (_paragraph.Format.SnapToGrid && _gridType != DocGridType.None)
            {
                double factor = 0.99;
                if (item is Character) factor = 0.90;
                while (item.Height > (_lineSpace * _rowCnt) * factor) _rowCnt++;
                _height = _lineSpace * _rowCnt;
            }
            else
            {
                _height = Math.Max(_height, item.Height);
            }
            // 调整行距
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
            // 设置底部内边距
            _maxCharHeight = Math.Max(_maxCharHeight, item.Height);
            _padding.Bottom = (_height - _maxCharHeight) / 2;
            // 如果分散对齐，在左右边距之间均匀分布文本, 起始字符左对齐，末尾字符右对齐，其余居中对齐
            if (_hAlign == HorizontalAlignment.Stretch)
            {
                if (_items.Count == 1) _items.First().HorizontalAlignment = HorizontalAlignment.Center;
                else
                {
                    int i = 0;
                    foreach (var c in _items)
                    {
                        if (i == 0) c.HorizontalAlignment = HorizontalAlignment.Left;
                        else if (i == _items.Count - 1) c.HorizontalAlignment = HorizontalAlignment.Right;
                        else c.HorizontalAlignment = HorizontalAlignment.Center;
                        i++;
                    }
                }
            }

            return true;
        }


        #endregion

    }
}
