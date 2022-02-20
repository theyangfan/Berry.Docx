using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent an OpenXML ParagraphProperties holder.
    /// </summary>
    internal class ParagraphPropertiesHolder
    {
        #region Private Members
        private Document _document = null;
        private W.ParagraphProperties _pPr = null;
        private W.StyleParagraphProperties _spPr = null;
        // Normal
        private W.Justification _justification = null;
        private W.OutlineLevel _outlineLevel = null;
        // Indentation
        private W.Indentation _indentation = null;
        private W.MirrorIndents _mirrorIndents = null;
        private W.AdjustRightIndent _adjustRightInd = null;
        // Spacing
        private W.SpacingBetweenLines _spacing = null;
        private W.SnapToGrid _snapToGrid = null;
        // Others
        private W.OverflowPunctuation _overflowPunct = null;
        private W.TopLinePunctuation _topLinePunct = null;
        // Numbering
        private W.Level _lvl = null;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the ParagraphPropertiesHolder class using the supplied <see cref="W.ParagraphProperties"/> element.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="pPr"></param>
        public ParagraphPropertiesHolder(Document document, W.ParagraphProperties pPr)
        {
            _document = document;
            if (pPr == null)
                pPr = new W.ParagraphProperties();
            _pPr = pPr;

            _justification = pPr.Justification;
            _outlineLevel = pPr.OutlineLevel;

            if (pPr.Indentation == null)
                pPr.Indentation = new W.Indentation();
            _indentation = pPr.Indentation;
            if (pPr.SpacingBetweenLines == null)
                pPr.SpacingBetweenLines = new W.SpacingBetweenLines();
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
                        W.Numbering numbering = _document.Package.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                        W.NumberingInstance num = numbering.Elements<W.NumberingInstance>().Where(n => n.NumberID == numId).FirstOrDefault();
                        if (num == null) return;
                        int abstractNumId = num.AbstractNumId.Val;
                        W.AbstractNum abstractNum = numbering.Elements<W.AbstractNum>().Where(a => a.AbstractNumberId == abstractNumId).FirstOrDefault();
                        if (abstractNum == null) return;
                        _lvl = abstractNum.Elements<W.Level>().Where(l => l.LevelIndex == ilvl).FirstOrDefault();
                    }
                }
            }
        }

        /// <summary>
        /// Initializes a new instance of the ParagraphPropertiesHolder class using the supplied OpenXML StyleParagraphProperties element.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="spPr"></param>
        public ParagraphPropertiesHolder(Document document, W.StyleParagraphProperties spPr)
        {
            _document = document;
            if (spPr == null)
                spPr = new W.StyleParagraphProperties();
            _spPr = spPr;
            _justification = spPr.Justification;
            _outlineLevel = spPr.OutlineLevel;

            if (spPr.Indentation == null)
                spPr.Indentation = new W.Indentation();
            _indentation = spPr.Indentation;
            if (spPr.SpacingBetweenLines == null)
                spPr.SpacingBetweenLines = new W.SpacingBetweenLines();
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
                    string styleId = (spPr.Parent as W.Style).StyleId;
                    if (_document.Package.MainDocumentPart.NumberingDefinitionsPart == null) return;
                    W.Numbering numbering = _document.Package.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                    W.NumberingInstance num = numbering.Elements<W.NumberingInstance>().Where(n => n.NumberID == numId).FirstOrDefault();
                    if (num == null) return;
                    int abstractNumId = num.AbstractNumId.Val;
                    W.AbstractNum abstractNum = numbering.Elements<W.AbstractNum>().Where(a => a.AbstractNumberId == abstractNumId).FirstOrDefault();
                    if (abstractNum == null) return;
                    _lvl = abstractNum.Elements<W.Level>().Where(l => l.ParagraphStyleIdInLevel != null && l.ParagraphStyleIdInLevel.Val == styleId).FirstOrDefault();
                }
            }
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets paragraph numbering format.
        /// </summary>
        public NumberingFormat NumberingFormat
        {
            get
            {
                if (_lvl == null) return null;
                return new NumberingFormat(_lvl);
            }
        }

        /// <summary>
        /// Gets or sets the justification.
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
                    _justification = new W.Justification();
                    if (_pPr != null)
                        _pPr.Justification = _justification;
                    else if (_spPr != null)
                        _spPr.Justification = _justification;
                }
                _justification.Val = value.Convert();
            }
        }

        /// <summary>
        /// Gets or sets the outline level.
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
                    _outlineLevel = new W.OutlineLevel();
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
        /// Gets or sets the left indent (in points) for paragraph.
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
        /// Gets or sets the right indent (in points) for paragraph.
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
        /// Gets or sets the left indent (in chars) for paragraph.
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
        /// Gets or sets the right indent (in chars) for paragraph.
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
        /// Gets or sets the first line indent (in points) for paragraph.
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
        /// Gets or sets the first line indent (in chars) for paragraph.
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
        /// Gets or sets the hanging indent (in points) for paragraph.
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
        /// Gets or sets the hanging indent (in chars) for paragraph.
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
        /// Gets or sets the spacing (in points) before the paragraph.
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
        /// Gets or sets the spacing (in lines) before the paragraph.
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
        /// Gets or sets a value indicating whether spacing before is automatic.
        /// </summary>
        public ZBool BeforeAutoSpacing
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
        /// Gets or sets the spacing (in points) after the paragraph.
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
        /// Gets or sets the spacing (in lines) after the paragraph.
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
        /// Gets or sets a value indicating whether spacing after is automatic.
        /// </summary>
        public ZBool AfterAutoSpacing
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
        /// Gets or sets the line spacing (in points) for paragraph.
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
        /// Gets or sets the line spacing rule of paragraph.
        /// </summary>
        public LineSpacingRule LineSpacingRule
        {
            get
            {
                if ( _spacing.LineRule == null) return LineSpacingRule.None;
                switch (_spacing.LineRule.Value)
                {
                    case W.LineSpacingRuleValues.Exact:
                        return LineSpacingRule.Exactly;
                    case W.LineSpacingRuleValues.AtLeast:
                        return LineSpacingRule.AtLeast;
                }
                return LineSpacingRule.Multiple;
            }
            set
            {
                switch (value)
                {
                    case LineSpacingRule.AtLeast:
                        _spacing.LineRule = W.LineSpacingRuleValues.AtLeast;
                        break;
                    case LineSpacingRule.Exactly:
                        _spacing.LineRule = W.LineSpacingRuleValues.Exact;
                        break;
                    case LineSpacingRule.Multiple:
                        _spacing.LineRule = W.LineSpacingRuleValues.Auto;
                        break;
                    case LineSpacingRule.None:
                        _spacing.LineRule = null;
                        break;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether allow punctuation to overflow boundaries.
        /// </summary>
        public ZBool OverflowPunctuation
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
                    _overflowPunct = new W.OverflowPunctuation();
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
        /// Gets or sets a value indicating whether allow top line punctuation compression.
        /// </summary>
        public ZBool TopLinePunctuation
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
                    _topLinePunct = new W.TopLinePunctuation();
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
        /// Gets or sets a value indicating whether the right indentation is automatically adjusted if a document grid is defined.
        /// </summary>
        public ZBool AdjustRightIndent
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
                    _adjustRightInd = new W.AdjustRightIndent();
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
        /// Gets or sets a value indicating whether snap to the grid if a document grid is defined.
        /// </summary>
        public ZBool SnapToGrid
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
                    _snapToGrid = new W.SnapToGrid();
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
        #endregion

        #region Public Methods
        /// <summary>
        /// Remove the text box options of paragraph.
        /// </summary>
        public void RemoveFrame()
        {
            if (_pPr != null)
                _pPr.FrameProperties = null;
            else if (_spPr != null)
                _spPr.FrameProperties = null;
        }
        #endregion
    }
}
