using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OOxml = DocumentFormat.OpenXml.Wordprocessing;
using P = DocumentFormat.OpenXml.Packaging;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// 段落格式
    /// </summary>
    public class ParagraphFormat
    {
        private Document _document = null;
        #region Private Paragraph Members
        private OOxml.Paragraph _ownerParagraph = null;
        private ParagraphPropertiesHolder _curPHld = null;
        private ParagraphFormat _stylePFormat = null;
        #endregion

        #region Private Style Members
        private OOxml.Style _ownerStyle = null;
        private ParagraphPropertiesHolder _curSHld = null;
        private ParagraphFormat _baseStylePFormat = null;
        #endregion

        #region Private Formats Menbers
        private JustificationType _justification = JustificationType.Both;
        private OutlineLevelType _outlineLevel = OutlineLevelType.BodyText;
        private float _leftIndent = -1;
        private float _rightIndent = -1;
        private float _leftCharsIndent = -1;
        private float _rightCharsIndent = -1;
        private float _firstLineIndent = -1;
        private float _firstLineCharsIndent = -1;
        private float _hangingIndent = -1;
        private float _hangingCharsIndent = -1;
        private float _beforeSpacing = -1;
        private float _beforeLinesSpacing = -1;
        private bool _beforeAutoSpacing = false;
        private float _afterSpacing = -1;
        private float _afterLinesSpacing = -1;
        private bool _afterAutoSpacing = false;
        private float _lineSpacing = 12;
        private LineSpacingRule _lineSpacingRule = LineSpacingRule.Multiple;
        private NumberingFormat _numFormat = null;

        private Zbool _overflowPunctuation = true;
        private Zbool _topLinePunctuation = false;
        private Zbool _adjustRightIndent = true;
        private Zbool _snapToGrid = true;
        #endregion

        /// <summary>
        /// 空构造函数
        /// </summary>
        public ParagraphFormat() { }
        /// <summary>
        /// 构造函数，用于构造段落的段落格式
        /// </summary>
        /// <param name="ownerParagraph"></param>
        public ParagraphFormat(Document document, OOxml.Paragraph ownerParagraph)
        {
            _document = document;
            _ownerParagraph = ownerParagraph;
            if (ownerParagraph.ParagraphProperties == null)
                ownerParagraph.ParagraphProperties = new OOxml.ParagraphProperties();
            _curPHld = new ParagraphPropertiesHolder(document, ownerParagraph.ParagraphProperties);
            _stylePFormat = new ParagraphFormat(document, ownerParagraph.GetStyle(document));
        }
        /// <summary>
        /// 构造函数，用于构造样式的段落格式
        /// </summary>
        /// <param name="ownerStyle"></param>
        public ParagraphFormat(Document document, OOxml.Style ownerStyle)
        {
            _document = document;
            _ownerStyle = ownerStyle;
            _curSHld = new ParagraphPropertiesHolder(document, ownerStyle.StyleParagraphProperties);
            _baseStylePFormat = GetStyleParagraphFormatRecursively(ownerStyle);
        }
        /// <summary>
        /// 递归地获取样式的段落格式
        /// </summary>
        /// <param name="style"></param>
        /// <returns></returns>
        private ParagraphFormat GetStyleParagraphFormatRecursively(OOxml.Style style)
        {
            ParagraphFormat format = new ParagraphFormat();
            ParagraphFormat baseFormat = new ParagraphFormat();
            OOxml.Style baseStyle = style.GetBaseStyle();
            if (baseStyle != null)
                baseFormat = GetStyleParagraphFormatRecursively(baseStyle);

            ParagraphPropertiesHolder curSHld = new ParagraphPropertiesHolder(_document, style.StyleParagraphProperties);

            format.NumberingFormat = curSHld.NumberingFormat ?? baseFormat.NumberingFormat;

            format.Justification = curSHld.Justification != JustificationType.None ? curSHld.Justification : baseFormat.Justification;
            format.OutlineLevel = curSHld.OutlineLevel != OutlineLevelType.None ? curSHld.OutlineLevel : baseFormat.OutlineLevel;
            format.LeftIndent = curSHld.LeftIndent >= 0 ? curSHld.LeftIndent : baseFormat.LeftIndent;
            format.LeftCharsIndent = curSHld.LeftCharsIndent >= 0 ? curSHld.LeftCharsIndent : baseFormat.LeftCharsIndent;
            format.RightIndent = curSHld.RightIndent >= 0 ? curSHld.RightIndent : baseFormat.RightIndent;
            format.RightCharsIndent = curSHld.RightCharsIndent >= 0 ? curSHld.RightCharsIndent : baseFormat.RightCharsIndent;
            format.FirstLineIndent = curSHld.FirstLineIndent >= 0 ? curSHld.FirstLineIndent : baseFormat.FirstLineIndent;
            format.FirstLineCharsIndent = curSHld.FirstLineCharsIndent >= 0 ? curSHld.FirstLineCharsIndent : baseFormat.FirstLineCharsIndent;
            format.HangingIndent = curSHld.HangingIndent >= 0 ? curSHld.HangingIndent : baseFormat.HangingIndent;
            format.HangingCharsIndent = curSHld.HangingCharsIndent >= 0 ? curSHld.HangingCharsIndent : baseFormat.HangingCharsIndent;
            format.BeforeSpacing = curSHld.BeforeSpacing >= 0 ? curSHld.BeforeSpacing : baseFormat.BeforeSpacing;
            format.BeforeLinesSpacing = curSHld.BeforeLinesSpacing >= 0 ? curSHld.BeforeLinesSpacing : baseFormat.BeforeLinesSpacing;
            format.BeforeAutoSpacing = curSHld.BeforeAutoSpacing ?? baseFormat.BeforeAutoSpacing;
            format.AfterSpacing = curSHld.AfterSpacing >= 0 ? curSHld.AfterSpacing : baseFormat.AfterSpacing;
            format.AfterLinesSpacing = curSHld.AfterLinesSpacing >= 0 ? curSHld.AfterLinesSpacing : baseFormat.AfterLinesSpacing;
            format.AfterAutoSpacing = curSHld.AfterAutoSpacing ?? baseFormat.AfterAutoSpacing;
            format.LineSpacing = curSHld.LineSpacing > 0 ? curSHld.LineSpacing : baseFormat.LineSpacing;
            format.LineSpacingRule = curSHld.LineSpacingRule != LineSpacingRule.None
                ? curSHld.LineSpacingRule : baseFormat.LineSpacingRule;
            format.OverflowPunctuation = curSHld.OverflowPunctuation ?? baseFormat.OverflowPunctuation;
            format.TopLinePunctuation = curSHld.TopLinePunctuation ?? baseFormat.TopLinePunctuation;
            format.AdjustRightIndent = curSHld.AdjustRightIndent ?? baseFormat.AdjustRightIndent;
            format.SnapToGrid = curSHld.SnapToGrid ?? baseFormat.SnapToGrid;
            return format;
        }

        /// <summary>
        /// 编号格式
        /// </summary>
        public NumberingFormat NumberingFormat
        {
            get
            {
                if(_ownerParagraph != null)
                {
                    return _ownerParagraph.ParagraphProperties != null && _ownerParagraph.ParagraphProperties.NumberingProperties != null ? _curPHld.NumberingFormat : _stylePFormat.NumberingFormat;
                }
                else if(_ownerStyle != null)
                {
                    return _baseStylePFormat.NumberingFormat;
                }
                else
                {
                    return _numFormat;
                }
            }
            set
            {
                if(_ownerParagraph == null && _ownerStyle == null)
                    _numFormat = value;
            }
        }
        /// <summary>
        /// 对齐方式
        /// </summary>
        public JustificationType Justification
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.Justification != JustificationType.None ? _curPHld.Justification : _stylePFormat.Justification;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.Justification;
                }
                else
                {
                    return _justification;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.Justification = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.Justification = value;
                }
                else
                {
                    _justification = value;
                }
            }
        }

        /// <summary>
        /// 大纲级别
        /// </summary>
        public OutlineLevelType OutlineLevel
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.OutlineLevel != OutlineLevelType.None ? _curPHld.OutlineLevel : _stylePFormat.OutlineLevel;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.OutlineLevel;
                }
                else
                {
                    return _outlineLevel;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.OutlineLevel = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.OutlineLevel = value;
                }
                else
                {
                    _outlineLevel = value;
                }
            }
        }

        /// <summary>
        /// 左侧缩进(磅)
        /// </summary>
        public float LeftIndent
        {
            get
            {
                if(_ownerParagraph != null)
                {
                    return _curPHld.LeftIndent >= 0 ? _curPHld.LeftIndent : _stylePFormat.LeftIndent;
                }
                else if(_ownerStyle != null)
                {
                    return _baseStylePFormat.LeftIndent;
                }
                else
                {
                    return _leftIndent;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.LeftIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.LeftIndent = value;
                }
                else
                {
                    _leftIndent = value;
                }
            }
        }
        /// <summary>
        /// 左侧缩进(字符)
        /// </summary>
        public float LeftCharsIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.LeftCharsIndent >= 0 ? _curPHld.LeftCharsIndent : _stylePFormat.LeftCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.LeftCharsIndent;
                }
                else
                {
                    return _leftCharsIndent;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.LeftCharsIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.LeftCharsIndent = value;
                }
                else
                {
                    _leftCharsIndent = value;
                }
            }
        }
        /// <summary>
        /// 右侧缩进(磅)
        /// </summary>
        public float RightIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.RightIndent >= 0 ? _curPHld.RightIndent : _stylePFormat.RightIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.RightIndent;
                }
                else
                {
                    return _rightIndent;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.RightIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.RightIndent = value;
                }
                else
                {
                    _rightIndent = value;
                }
            }
        }
        /// <summary>
        /// 右侧缩进(字符)
        /// </summary>
        public float RightCharsIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.RightCharsIndent >= 0 ? _curPHld.RightCharsIndent : _stylePFormat.RightCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.RightCharsIndent;
                }
                else
                {
                    return _rightCharsIndent;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.RightCharsIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.RightCharsIndent = value;
                }
                else
                {
                    _rightCharsIndent = value;
                }
            }
        }
        /// <summary>
        /// 首行缩进(磅)
        /// </summary>
        public float FirstLineIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.FirstLineIndent >= 0 ? _curPHld.FirstLineIndent : _stylePFormat.FirstLineIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.FirstLineIndent;
                }
                else
                {
                    return _firstLineIndent;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.FirstLineIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.FirstLineIndent = value;
                }
                else
                {
                    _firstLineIndent = value;
                }
            }
        }
        /// <summary>
        /// 首行缩进(字符)
        /// </summary>
        public float FirstLineCharsIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.FirstLineCharsIndent >= 0 ? _curPHld.FirstLineCharsIndent : _stylePFormat.FirstLineCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.FirstLineCharsIndent;
                }
                else
                {
                    return _firstLineCharsIndent;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.FirstLineCharsIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.FirstLineCharsIndent = value;
                }
                else
                {
                    _firstLineCharsIndent = value;
                }
            }
        }
        /// <summary>
        /// 悬挂缩进(磅)
        /// </summary>
        public float HangingIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.HangingIndent >= 0 ? _curPHld.HangingIndent : _stylePFormat.HangingIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.HangingIndent;
                }
                else
                {
                    return _hangingIndent;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.HangingIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.HangingIndent = value;
                }
                else
                {
                    _hangingIndent = value;
                }
            }
        }
        /// <summary>
        /// 悬挂缩进(字符)
        /// </summary>
        public float HangingCharsIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.HangingCharsIndent >= 0 ? _curPHld.HangingCharsIndent : _stylePFormat.HangingCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.HangingCharsIndent;
                }
                else
                {
                    return _hangingCharsIndent;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.HangingCharsIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.HangingCharsIndent = value;
                }
                else
                {
                    _hangingCharsIndent = value;
                }
            }
        }
        /// <summary>
        /// 段前间距(磅)
        /// </summary>
        public float BeforeSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.BeforeSpacing >= 0 ? _curPHld.BeforeSpacing : _stylePFormat.BeforeSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.BeforeSpacing;
                }
                else
                {
                    return _beforeSpacing;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.BeforeSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.BeforeSpacing = value;
                }
                else
                {
                    _beforeSpacing = value;
                }
            }
        }
        /// <summary>
        /// 段前间距(行)
        /// </summary>
        public float BeforeLinesSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.BeforeLinesSpacing >= 0 ? _curPHld.BeforeLinesSpacing : _stylePFormat.BeforeLinesSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.BeforeLinesSpacing;
                }
                else
                {
                    return _beforeLinesSpacing;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.BeforeLinesSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.BeforeLinesSpacing = value;
                }
                else
                {
                    _beforeLinesSpacing = value;
                }
            }
        }
        /// <summary>
        /// 自动调整段前间距
        /// </summary>
        public bool BeforeAutoSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.BeforeAutoSpacing ?? _stylePFormat.BeforeAutoSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.BeforeAutoSpacing;
                }
                else
                {
                    return _beforeAutoSpacing;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.BeforeAutoSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.BeforeAutoSpacing = value;
                }
                else
                {
                    _beforeAutoSpacing = value;
                }
            }
        }
        /// <summary>
        /// 段后间距(磅)
        /// </summary>
        public float AfterSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.AfterSpacing >= 0 ? _curPHld.AfterSpacing : _stylePFormat.AfterSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.AfterSpacing;
                }
                else
                {
                    return _afterSpacing;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.AfterSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.AfterSpacing = value;
                }
                else
                {
                    _afterSpacing = value;
                }
            }
        }
        /// <summary>
        /// 段后间距(行)
        /// </summary>
        public float AfterLinesSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.AfterLinesSpacing >= 0 ? _curPHld.AfterLinesSpacing : _stylePFormat.AfterLinesSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.AfterLinesSpacing;
                }
                else
                {
                    return _afterLinesSpacing;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.AfterLinesSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.AfterLinesSpacing = value;
                }
                else
                {
                    _afterLinesSpacing = value;
                }
            }
        }

        /// <summary>
        /// 自动调整段后间距
        /// </summary>
        public bool AfterAutoSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.AfterAutoSpacing ?? _stylePFormat.AfterAutoSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.AfterAutoSpacing;
                }
                else
                {
                    return _afterAutoSpacing;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.AfterAutoSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.AfterAutoSpacing = value;
                }
                else
                {
                    _afterAutoSpacing = value;
                }
            }
        }
        /// <summary>
        /// 行距(磅)
        /// </summary>
        public float LineSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.LineSpacing >= 0 ? _curPHld.LineSpacing : _stylePFormat.LineSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.LineSpacing;
                }
                else
                {
                    return _lineSpacing;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.LineSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.LineSpacing = value;
                }
                else
                {
                    _lineSpacing = value;
                }
            }
        }
        /// <summary>
        /// 行距类型
        /// </summary>
        public LineSpacingRule LineSpacingRule
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.LineSpacingRule != LineSpacingRule.None ? _curPHld.LineSpacingRule : _stylePFormat.LineSpacingRule;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.LineSpacingRule;
                }
                else
                {
                    return _lineSpacingRule;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.LineSpacingRule = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.LineSpacingRule = value;
                }
                else
                {
                    _lineSpacingRule = value;
                }
            }
        }

        /// <summary>
        /// 允许标点溢出边界
        /// </summary>
        public bool OverflowPunctuation
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.OverflowPunctuation ?? _stylePFormat.OverflowPunctuation;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.OverflowPunctuation;
                }
                else
                {
                    return _overflowPunctuation;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.OverflowPunctuation = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.OverflowPunctuation = value;
                }
                else
                {
                    _overflowPunctuation = value;
                }
            }
        }

        /// <summary>
        /// 允许行首标点压缩
        /// </summary>
        public bool TopLinePunctuation
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.TopLinePunctuation ?? _stylePFormat.TopLinePunctuation;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.TopLinePunctuation;
                }
                else
                {
                    return _topLinePunctuation;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.TopLinePunctuation = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.TopLinePunctuation = value;
                }
                else
                {
                    _topLinePunctuation = value;
                }
            }
        }
        /// <summary>
        /// 如果定义了文档网格，则自动调整右缩进
        /// </summary>
        public bool AdjustRightIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.AdjustRightIndent ?? _stylePFormat.AdjustRightIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.AdjustRightIndent;
                }
                else
                {
                    return _adjustRightIndent;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.AdjustRightIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.AdjustRightIndent = value;
                }
                else
                {
                    _adjustRightIndent = value;
                }
            }
        }

        /// <summary>
        /// 如果定义了文档网格，则对齐到网格
        /// </summary>
        public bool SnapToGrid
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.SnapToGrid ?? _stylePFormat.SnapToGrid;
                }
                else if (_ownerStyle != null)
                {
                    return _baseStylePFormat.SnapToGrid;
                }
                else
                {
                    return _snapToGrid;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.SnapToGrid = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.SnapToGrid = value;
                }
                else
                {
                    _snapToGrid = value;
                }
            }
        }

        /// <summary>
        /// 去除文本框选项
        /// </summary>
        public void RemoveFrame()
        {
            if (_ownerParagraph != null)
            {
                _curPHld.RemoveFrame();
            }
            else if (_ownerStyle != null)
            {
                _curSHld.RemoveFrame();
            }
        }

    }
}
