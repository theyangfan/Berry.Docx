using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OOxml = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    internal class ParagraphPropertiesHolder
    {
        private Document _document = null;

        private OOxml.ParagraphProperties _pPr = null;
        private OOxml.StyleParagraphProperties _spPr = null;
        // 常规
        private OOxml.Justification _justification = null;
        private OOxml.OutlineLevel _outlineLevel = null;
        // 缩进
        private OOxml.Indentation _indentation = null;
        private OOxml.MirrorIndents _mirrorIndents = null;
        private OOxml.AdjustRightIndent _adjustRightInd = null;
        // 间距
        private OOxml.SpacingBetweenLines _spacing = null;
        private OOxml.SnapToGrid _snapToGrid = null;

        private OOxml.OverflowPunctuation _overflowPunct = null;
        private OOxml.TopLinePunctuation _topLinePunct = null;
        // 编号
        private OOxml.Level _lvl = null;

        public ParagraphPropertiesHolder(Document document, OOxml.ParagraphProperties pPr)
        {
            _document = document;
            if (pPr == null)
                pPr = new OOxml.ParagraphProperties();
            _pPr = pPr;

            _justification = pPr.Justification;
            _outlineLevel = pPr.OutlineLevel;

            if (pPr.Indentation == null)
                pPr.Indentation = new OOxml.Indentation();
            _indentation = pPr.Indentation;
            if (pPr.SpacingBetweenLines == null)
                pPr.SpacingBetweenLines = new OOxml.SpacingBetweenLines();
            _spacing = pPr.SpacingBetweenLines;

            _overflowPunct = pPr.OverflowPunctuation;
            _topLinePunct = pPr.TopLinePunctuation;
            _adjustRightInd = pPr.AdjustRightIndent;
            _snapToGrid = pPr.SnapToGrid;
            if(pPr.NumberingProperties != null)
            {
                if(pPr.NumberingProperties.NumberingId != null)
                {
                    int numId = pPr.NumberingProperties.NumberingId.Val;
                    if(pPr.NumberingProperties.NumberingLevelReference != null)
                    {
                        int ilvl = pPr.NumberingProperties.NumberingLevelReference.Val;
                        if (_document.Package.MainDocumentPart.NumberingDefinitionsPart == null) return;
                        OOxml.Numbering numbering = _document.Package.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                        OOxml.NumberingInstance num = numbering.Elements<OOxml.NumberingInstance>().Where(n => n.NumberID == numId).FirstOrDefault();
                        if (num == null) return;
                        int abstractNumId = num.AbstractNumId.Val;
                        OOxml.AbstractNum abstractNum = numbering.Elements<OOxml.AbstractNum>().Where(a => a.AbstractNumberId == abstractNumId).FirstOrDefault();
                        if (abstractNum == null) return;
                        _lvl = abstractNum.Elements<OOxml.Level>().Where(l => l.LevelIndex == ilvl).FirstOrDefault();
                    }
                }
            }
        }

        public ParagraphPropertiesHolder(Document document, OOxml.StyleParagraphProperties spPr)
        {
            _document = document;
            if (spPr == null)
                spPr = new OOxml.StyleParagraphProperties();
            _spPr = spPr;
            _justification = spPr.Justification;
            _outlineLevel = spPr.OutlineLevel;

            if (spPr.Indentation == null)
                spPr.Indentation = new OOxml.Indentation();
            _indentation = spPr.Indentation;
            if (spPr.SpacingBetweenLines == null)
                spPr.SpacingBetweenLines = new OOxml.SpacingBetweenLines();
            _spacing = spPr.SpacingBetweenLines;

            _overflowPunct = spPr.OverflowPunctuation;
            _topLinePunct = spPr.TopLinePunctuation;
            _adjustRightInd = spPr.AdjustRightIndent;
            _snapToGrid = spPr.SnapToGrid;
            if(spPr.NumberingProperties != null)
            {
                if (spPr.NumberingProperties.NumberingId != null)
                {
                    int numId = spPr.NumberingProperties.NumberingId.Val;
                    string styleId = (spPr.Parent as OOxml.Style).StyleId;
                    if (_document.Package.MainDocumentPart.NumberingDefinitionsPart == null) return;
                    OOxml.Numbering numbering = _document.Package.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                    OOxml.NumberingInstance num = numbering.Elements<OOxml.NumberingInstance>().Where(n => n.NumberID == numId).FirstOrDefault();
                    if (num == null) return;
                    int abstractNumId = num.AbstractNumId.Val;
                    OOxml.AbstractNum abstractNum = numbering.Elements<OOxml.AbstractNum>().Where(a => a.AbstractNumberId == abstractNumId).FirstOrDefault();
                    if (abstractNum == null) return;
                    _lvl = abstractNum.Elements<OOxml.Level>().Where(l => l.ParagraphStyleIdInLevel != null && l.ParagraphStyleIdInLevel.Val == styleId).FirstOrDefault();
                }
            }
        }

        public NumberingFormat NumberingFormat
        {
            get
            {
                if (_lvl == null) return null;
                return new NumberingFormat(_lvl);
            }
        }
        /// <summary>
        /// 对齐方式
        /// </summary>
        public JustificationType Justification
        {
            get
            {
                if (_justification == null) return JustificationType.None;
                return _justification.Val.Value.Convert();
            }
            set
            {
                if (_justification == null)
                {
                    _justification = new OOxml.Justification();
                    if (_pPr != null)
                        _pPr.Justification = _justification;
                    else if (_spPr != null)
                        _spPr.Justification = _justification;
                }
                _justification.Val = value.Convert();
            }
        }

        /// <summary>
        /// 大纲级别
        /// </summary>
        public OutlineLevelType OutlineLevel
        {
            get
            {
                if (_outlineLevel == null) return OutlineLevelType.None;
                return (OutlineLevelType)_outlineLevel.Val.Value;
            }
            set
            {
                if (_outlineLevel == null)
                {
                    _outlineLevel = new OOxml.OutlineLevel();
                    if (_pPr != null)
                        _pPr.OutlineLevel = _outlineLevel;
                    else if (_spPr != null)
                        _spPr.OutlineLevel = _outlineLevel;
                }
                if (value != OutlineLevelType.BodyText)
                    _outlineLevel.Val = (int)value;
                else
                    _outlineLevel = null;
            }
        }

        /// <summary>
        /// 左侧缩进(磅)
        /// </summary>
        public float LeftIndent
        {
            get
            {
                float val = 0;
                if (_indentation.Left == null) return -1;
                float.TryParse(_indentation.Left, out val);
                val = val / 20;
                if (HangingCharsIndent > 0)
                    val = 0;
                else if (HangingIndent > 0)
                    val -= HangingIndent;
                return val;
            }
            set
            {
                if (value >= 0)
                {
                    if (HangingIndent > 0)
                        _indentation.Left = ((value + HangingIndent) * 20).ToString();
                    else
                        _indentation.Left = (value * 20).ToString();
                }
                else
                    _indentation.Left = null;
            }
        }

        /// <summary>
        /// 右侧缩进(磅)
        /// </summary>
        
        public float RightIndent
        {
            get
            {
                float val = 0;
                if (_indentation.Right == null) return -1;
                float.TryParse(_indentation.Right, out val);
                return val / 20;
            }
            set
            {
                if (value >= 0)
                    _indentation.Right = (value * 20).ToString();
                else
                    _indentation.Right = null;
            }
        }
        
        /// <summary>
        /// 左侧缩进(字符)
        /// </summary>
        public float LeftCharsIndent
        {
            get
            {
                float val = 0;
                if (_indentation.LeftChars == null) return -1;
                float.TryParse(_indentation.LeftChars, out val);
                return val / 100;
            }
            set
            {
                if (value >= 0)
                    _indentation.LeftChars = (int)(value * 100);
                else
                    _indentation.LeftChars = null;
            }
        }
        /// <summary>
        /// 右侧缩进(字符)
        /// </summary>
        public float RightCharsIndent
        {
            get
            {
                float val = 0;
                if (_indentation.RightChars == null) return -1;
                float.TryParse(_indentation.RightChars, out val);
                return val / 100;
            }
            set
            {
                if (value >= 0)
                    _indentation.RightChars = (int)(value * 100);
                else
                    _indentation.RightChars = null;
            }
        }
        /// <summary>
        /// 首行缩进(磅)
        /// </summary>
        public float FirstLineIndent
        {
            get
            {
                float val = 0;
                if (_indentation.FirstLine == null) return -1;
                float.TryParse(_indentation.FirstLine, out val);
                return val / 20;
            }
            set
            {
                if (value >= 0)
                {
                    HangingIndent = -1;
                    HangingCharsIndent = -1;
                    _indentation.FirstLine = (value * 20).ToString();
                }
                else
                {
                    _indentation.FirstLine = null;
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
                float val = 0;
                if (_indentation.FirstLineChars == null) return -1;
                float.TryParse(_indentation.FirstLineChars, out val);
                return val / 100;
            }
            set
            {
                if (value >= 0)
                {
                    HangingIndent = -1;
                    HangingCharsIndent = -1;
                    _indentation.FirstLineChars = (int)(value * 100);
                }
                else
                {
                    _indentation.FirstLineChars = null;
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
                float val = 0;
                if (_indentation.Hanging == null) return -1;
                float.TryParse(_indentation.Hanging, out val);
                return val / 20;
            }
            set
            {
                if (value >= 0)
                {
                    FirstLineIndent = -1;
                    FirstLineCharsIndent = -1;
                    _indentation.Hanging = (value * 20).ToString();
                }
                else
                {
                    _indentation.Hanging = null;
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
                float val = 0;
                if (_indentation.HangingChars == null) return -1;
                float.TryParse(_indentation.HangingChars, out val);
                return val / 100;
            }
            set
            {
                if (value >= 0)
                {
                    FirstLineIndent = -1;
                    FirstLineCharsIndent = -1;
                    _indentation.HangingChars = (int)(value * 100);
                }
                else
                {
                    _indentation.HangingChars = null;
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
                float val = 0;
                if (_spacing.Before == null) return -1;
                float.TryParse(_spacing.Before, out val);
                return val / 20;
            }
            set
            {
                if (value >= 0)
                {
                    _spacing.Before = (value * 20).ToString();
                }
                else
                {
                    _spacing.Before = null;
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
                float val = 0;
                if (_spacing.BeforeLines == null) return -1;
                float.TryParse(_spacing.BeforeLines, out val);
                return val / 100;
            }
            set
            {
                if (value >= 0)
                {
                    _spacing.BeforeLines = (int)(value * 100);
                }
                else
                {
                    _spacing.BeforeLines = null;
                }
            }
        }

        /// <summary>
        /// 自动调整段前间距
        /// </summary>
        public Zbool BeforeAutoSpacing
        {
            get
            {
                if (_spacing.BeforeAutoSpacing == null) return null;
                return _spacing.BeforeAutoSpacing.Value;
            }
            set
            {
                if (value)
                    _spacing.BeforeAutoSpacing = true;
                else
                    _spacing.BeforeAutoSpacing = false;
            }
        }

        /// <summary>
        /// 段后间距(磅)
        /// </summary>
        public float AfterSpacing
        {
            get
            {
                float val = 0;
                if (_spacing.After == null) return -1;
                float.TryParse(_spacing.After, out val);
                return val / 20;
            }
            set
            {
                if (value >= 0)
                {
                    _spacing.After = (value * 20).ToString();
                }
                else
                {
                    _spacing.After = null;
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
                float val = 0;
                if (_spacing.AfterLines == null) return -1;
                float.TryParse(_spacing.AfterLines, out val);
                return val / 100;
            }
            set
            {
                if (value >= 0)
                {
                    _spacing.AfterLines = (int)(value * 100);
                }
                else
                {
                    _spacing.AfterLines = null;
                }
            }
        }

        /// <summary>
        /// 自动调整段后间距
        /// </summary>
        public Zbool AfterAutoSpacing
        {
            get
            {
                if (_spacing.AfterAutoSpacing == null) return null;
                return _spacing.AfterAutoSpacing.Value;
            }
            set
            {
                if (value)
                    _spacing.AfterAutoSpacing = true;
                else
                    _spacing.AfterAutoSpacing = false;
            }
        }

        /// <summary>
        /// 行距(磅)
        /// </summary>
        public float LineSpacing
        {
            get
            {
                float val = 0;
                if (_spacing.Line == null) return -1;
                float.TryParse(_spacing.Line, out val);
                return val / 20;
            }
            set
            {
                if (value >= 0)
                {
                    _spacing.Line = (value * 20).ToString();
                }
                else
                {
                    _spacing.Line = null;
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
                if ( _spacing.LineRule == null) return LineSpacingRule.None;
                switch (_spacing.LineRule.Value)
                {
                    case OOxml.LineSpacingRuleValues.Exact:
                        return LineSpacingRule.Exactly;
                    case OOxml.LineSpacingRuleValues.AtLeast:
                        return LineSpacingRule.AtLeast;
                }
                return LineSpacingRule.Multiple;
            }
            set
            {
                switch (value)
                {
                    case LineSpacingRule.AtLeast:
                        _spacing.LineRule = OOxml.LineSpacingRuleValues.AtLeast;
                        break;
                    case LineSpacingRule.Exactly:
                        _spacing.LineRule = OOxml.LineSpacingRuleValues.Exact;
                        break;
                    case LineSpacingRule.Multiple:
                        _spacing.LineRule = OOxml.LineSpacingRuleValues.Auto;
                        break;
                    case LineSpacingRule.None:
                        _spacing.LineRule = null;
                        break;
                }
            }
        }

        /// <summary>
        /// 允许标点溢出边界
        /// </summary>
        public Zbool OverflowPunctuation
        {
            get
            {
                if (_overflowPunct == null) return null;
                if (_overflowPunct.Val == null) return true;
                return _overflowPunct.Val.Value;
            }
            set
            {
                if (_overflowPunct == null)
                {
                    _overflowPunct = new OOxml.OverflowPunctuation();
                    if (_pPr != null)
                        _pPr.OverflowPunctuation = _overflowPunct;
                    else if (_spPr != null)
                        _spPr.OverflowPunctuation = _overflowPunct;
                }
                if (value)
                    _overflowPunct.Val = null;
                else
                    _overflowPunct.Val = false;
            }
        }
        /// <summary>
        /// 允许行首标点压缩
        /// </summary>
        public Zbool TopLinePunctuation
        {
            get
            {
                if (_topLinePunct == null) return null;
                if (_topLinePunct.Val == null) return true;
                return _topLinePunct.Val.Value;
            }
            set
            {
                if (_topLinePunct == null)
                {
                    _topLinePunct = new OOxml.TopLinePunctuation();
                    if (_pPr != null)
                        _pPr.TopLinePunctuation = _topLinePunct;
                    else if (_spPr != null)
                        _spPr.TopLinePunctuation = _topLinePunct;
                }
                if (value)
                    _topLinePunct.Val = null;
                else
                    _topLinePunct.Val = false;
            }
        }

        /// <summary>
        /// 如果定义了文档网格，则自动调整右缩进
        /// </summary>
        public Zbool AdjustRightIndent
        {
            get
            {
                if (_adjustRightInd == null) return null;
                if (_adjustRightInd.Val == null) return true;
                return _adjustRightInd.Val.Value;
            }
            set
            {
                if (_adjustRightInd == null)
                {
                    _adjustRightInd = new OOxml.AdjustRightIndent();
                    if(_pPr != null)
                        _pPr.AdjustRightIndent = _adjustRightInd;
                    else if(_spPr != null)
                        _spPr.AdjustRightIndent = _adjustRightInd;
                }
                    
                if (value)
                    _adjustRightInd.Val = null;
                else
                    _adjustRightInd.Val = false;
            }
        }

        /// <summary>
        /// 如果定义了文档网格，则对齐到网格
        /// </summary>
        public Zbool SnapToGrid
        {
            get
            {
                if (_snapToGrid == null) return null;
                if (_snapToGrid.Val == null) return true;
                return _snapToGrid.Val.Value;
            }
            set
            {
                if (_snapToGrid == null)
                {
                    _snapToGrid = new OOxml.SnapToGrid();
                    if (_pPr != null)
                        _pPr.SnapToGrid = _snapToGrid;
                    else if (_spPr != null)
                        _spPr.SnapToGrid = _snapToGrid;
                }
                if (value)
                    _snapToGrid.Val = null;
                else
                    _snapToGrid.Val = false;
            }
        }
        /// <summary>
        /// 去除文本框选项
        /// </summary>
        public void RemoveFrame()
        {
            if (_pPr != null)
                _pPr.FrameProperties = null;
            else if (_spPr != null)
                _spPr.FrameProperties = null;
        }

    }
}
