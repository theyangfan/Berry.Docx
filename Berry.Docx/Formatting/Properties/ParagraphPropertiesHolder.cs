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
        private W.ContextualSpacing _contextualSpacing = null;
        private W.SnapToGrid _snapToGrid = null;
        // Pagination
        private W.WidowControl _widowControl;
        private W.KeepNext _keepNext;
        private W.KeepLines _keepLines;
        private W.PageBreakBefore _pageBreakBefore;
        // Formatting Exceptions
        private W.SuppressLineNumbers _suppressLineNumbers;
        private W.SuppressAutoHyphens _suppressAutoHyphens;
        // Line Break
        private W.Kinsoku _kinsoku;
        private W.WordWrap _wordWrap;
        private W.OverflowPunctuation _overflowPunct = null;
        // Character Spacing
        private W.TopLinePunctuation _topLinePunct = null;
        private W.AutoSpaceDE _autoSpaceDE;
        private W.AutoSpaceDN _autoSpaceDN;
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
            // Normal
            _justification = pPr.Justification;
            _outlineLevel = pPr.OutlineLevel;
            // Indentation
            if (pPr.Indentation == null)
                pPr.Indentation = new W.Indentation();
            _indentation = pPr.Indentation;
            _mirrorIndents = pPr.MirrorIndents;
            _adjustRightInd = pPr.AdjustRightIndent;
            // Spacing
            if (pPr.SpacingBetweenLines == null)
                pPr.SpacingBetweenLines = new W.SpacingBetweenLines();
            _spacing = pPr.SpacingBetweenLines;
            _contextualSpacing = pPr.ContextualSpacing;
            _snapToGrid = pPr.SnapToGrid;
            // Pagination
            _widowControl = pPr.WidowControl;
            _keepNext = pPr.KeepNext;
            _keepLines = pPr.KeepLines;
            _pageBreakBefore = pPr.PageBreakBefore;
            // Format Exception
            _suppressLineNumbers = pPr.SuppressLineNumbers;
            _suppressAutoHyphens = pPr.SuppressAutoHyphens;
            // Wrapping Lines
            _kinsoku = pPr.Kinsoku;
            _wordWrap = pPr.WordWrap;
            _overflowPunct = pPr.OverflowPunctuation;
            // Character Spacing
            _topLinePunct = pPr.TopLinePunctuation;
            _autoSpaceDE = pPr.AutoSpaceDE;
            _autoSpaceDN = pPr.AutoSpaceDN;
            // Numbering
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
            // Normal
            _justification = spPr.Justification;
            _outlineLevel = spPr.OutlineLevel;
            // Indentation
            if (spPr.Indentation == null)
                spPr.Indentation = new W.Indentation();
            _indentation = spPr.Indentation;
            _mirrorIndents = spPr.MirrorIndents;
            _adjustRightInd = spPr.AdjustRightIndent;
            // Spacing
            if (spPr.SpacingBetweenLines == null)
                spPr.SpacingBetweenLines = new W.SpacingBetweenLines();
            _spacing = spPr.SpacingBetweenLines;
            _contextualSpacing = spPr.ContextualSpacing;
            _snapToGrid = spPr.SnapToGrid;
            // Pagination
            _widowControl = spPr.WidowControl;
            _keepNext = spPr.KeepNext;
            _keepLines = spPr.KeepLines;
            _pageBreakBefore = spPr.PageBreakBefore;
            // Format Exception
            _suppressLineNumbers = spPr.SuppressLineNumbers;
            _suppressAutoHyphens = spPr.SuppressAutoHyphens;
            // Wrapping Lines
            _kinsoku = spPr.Kinsoku;
            _wordWrap = spPr.WordWrap;
            _overflowPunct = spPr.OverflowPunctuation;
            // Character Spacing
            _topLinePunct = spPr.TopLinePunctuation;
            _autoSpaceDE = spPr.AutoSpaceDE;
            _autoSpaceDN = spPr.AutoSpaceDN;
            // Numbering
            if (spPr.NumberingProperties != null)
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

        #region Normal
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
        #endregion

        #region Indentation
        /// <summary>
        /// Gets or sets the left indent (in points) for paragraph.
        /// </summary>
        public FloatValue LeftIndent
        {
            get
            {
                float val = 0;
                if (_indentation.Left == null) return null;
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

        public FloatValue RightIndent
        {
            get
            {
                float val = 0;
                if (_indentation.Right == null) return null;
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
        public FloatValue LeftCharsIndent
        {
            get
            {
                float val = 0;
                if (_indentation.LeftChars == null) return null;
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
        public FloatValue RightCharsIndent
        {
            get
            {
                float val = 0;
                if (_indentation.RightChars == null) return null;
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
        public FloatValue FirstLineIndent
        {
            get
            {
                float val = 0;
                if (_indentation.FirstLine == null) return null;
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
        public FloatValue FirstLineCharsIndent
        {
            get
            {
                float val = 0;
                if (_indentation.FirstLineChars == null) return null;
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
        public FloatValue HangingIndent
        {
            get
            {
                float val = 0;
                if (_indentation.Hanging == null) return null;
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
        public FloatValue HangingCharsIndent
        {
            get
            {
                float val = 0;
                if (_indentation.HangingChars == null) return null;
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
        /// Gets or sets a value indicating whether automatically adjust right indent when document grid is defined.
        /// </summary>
        public BooleanValue AdjustRightIndent
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
                    if (_pPr != null)
                        _pPr.AdjustRightIndent = _adjustRightInd;
                    else if (_spPr != null)
                        _spPr.AdjustRightIndent = _adjustRightInd;
                }

                if (value)
                    _adjustRightInd.Val = null;
                else
                    _adjustRightInd.Val = false;
            }
        }
        /// <summary>
        /// Gets or sets a value indicating whether the paragraph indents should be interpreted as mirrored indents.
        /// </summary>
        public BooleanValue MirrorIndents
        {
            get
            {
                if (_mirrorIndents == null) return null;
                if (_mirrorIndents.Val == null) return true;
                return _mirrorIndents.Val.Value;
            }
            set
            {
                if (_mirrorIndents == null)
                {
                    _mirrorIndents = new W.MirrorIndents();
                    if (_pPr != null)
                        _pPr.MirrorIndents = _mirrorIndents;
                    else if (_spPr != null)
                        _spPr.MirrorIndents = _mirrorIndents;
                }

                if (value)
                    _mirrorIndents.Val = null;
                else
                    _mirrorIndents.Val = false;
            }
        }
        #endregion

        #region Spacing
        /// <summary>
        /// Gets or sets the spacing (in points) before the paragraph.
        /// </summary>
        public FloatValue BeforeSpacing
        {
            get
            {
                float val = 0;
                if (_spacing.Before == null) return null;
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
        public FloatValue BeforeLinesSpacing
        {
            get
            {
                float val = 0;
                if (_spacing.BeforeLines == null) return null;
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
        public BooleanValue BeforeAutoSpacing
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
        public FloatValue AfterSpacing
        {
            get
            {
                float val = 0;
                if (_spacing.After == null) return null;
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
        public FloatValue AfterLinesSpacing
        {
            get
            {
                float val = 0;
                if (_spacing.AfterLines == null) return null;
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
        public BooleanValue AfterAutoSpacing
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
        public FloatValue LineSpacing
        {
            get
            {
                float val = 0;
                if (_spacing.Line == null) return null;
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
        /// Gets or sets a value indicating whether don't add space between paragraphs of the same style.
        /// </summary>
        public BooleanValue ContextualSpacing
        {
            get
            {
                if (_contextualSpacing == null) return null;
                if (_contextualSpacing.Val == null) return true;
                return _contextualSpacing.Val.Value;
            }
            set
            {
                if (_contextualSpacing == null)
                {
                    _contextualSpacing = new W.ContextualSpacing();
                    if (_pPr != null)
                        _pPr.ContextualSpacing = _contextualSpacing;
                    else if (_spPr != null)
                        _spPr.ContextualSpacing = _contextualSpacing;
                }
                if (value)
                    _contextualSpacing.Val = null;
                else
                    _contextualSpacing.Val = false;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether snap to grid when document grid is defined.
        /// </summary>
        public BooleanValue SnapToGrid
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

        #region Pagination
        /// <summary>
        /// Gets or sets a value indicating whether a consumer shall prevent first/last line of this paragraph 
        /// from being displayed on a separate page.
        /// </summary>
        public BooleanValue WidowControl
        {
            get
            {
                if (_widowControl == null) return null;
                if (_widowControl.Val == null) return true;
                return _widowControl.Val.Value;
            }
            set
            {
                if (_widowControl == null)
                {
                    _widowControl = new W.WidowControl();
                    if (_pPr != null)
                        _pPr.WidowControl = _widowControl;
                    else if (_spPr != null)
                        _spPr.WidowControl = _widowControl;
                }
                if (value)
                    _widowControl.Val = null;
                else
                    _widowControl.Val = false;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether keep paragraph with next paragraph on the same page.
        /// </summary>
        public BooleanValue KeepNext
        {
            get
            {
                if (_keepNext == null) return null;
                if (_keepNext.Val == null) return true;
                return _keepNext.Val.Value;
            }
            set
            {
                if (_keepNext == null)
                {
                    _keepNext = new W.KeepNext();
                    if (_pPr != null)
                        _pPr.KeepNext = _keepNext;
                    else if (_spPr != null)
                        _spPr.KeepNext = _keepNext;
                }
                if (value)
                    _keepNext.Val = null;
                else
                    _keepNext.Val = false;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether keep all lines of this paragraph on one page.
        /// </summary>
        public BooleanValue KeepLines
        {
            get
            {
                if (_keepLines == null) return null;
                if (_keepLines.Val == null) return true;
                return _keepLines.Val.Value;
            }
            set
            {
                if (_keepLines == null)
                {
                    _keepLines = new W.KeepLines();
                    if (_pPr != null)
                        _pPr.KeepLines = _keepLines;
                    else if (_spPr != null)
                        _spPr.KeepLines = _keepLines;
                }
                if (value)
                    _keepLines.Val = null;
                else
                    _keepLines.Val = false;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether start paragraph on next page.
        /// </summary>
        public BooleanValue PageBreakBefore
        {
            get
            {
                if (_pageBreakBefore == null) return null;
                if (_pageBreakBefore.Val == null) return true;
                return _pageBreakBefore.Val.Value;
            }
            set
            {
                if (_pageBreakBefore == null)
                {
                    _pageBreakBefore = new W.PageBreakBefore();
                    if (_pPr != null)
                        _pPr.PageBreakBefore = _pageBreakBefore;
                    else if (_spPr != null)
                        _spPr.PageBreakBefore = _pageBreakBefore;
                }
                if (value)
                    _pageBreakBefore.Val = null;
                else
                    _pageBreakBefore.Val = false;
            }
        }
        #endregion

        #region Formatting Exceptions
        /// <summary>
        /// Gets or sets a value indicating whether suppress line numbers for paragraph.
        /// </summary>
        public BooleanValue SuppressLineNumbers
        {
            get
            {
                if (_suppressLineNumbers == null) return null;
                if (_suppressLineNumbers.Val == null) return true;
                return _suppressLineNumbers.Val.Value;
            }
            set
            {
                if (_suppressLineNumbers == null)
                {
                    _suppressLineNumbers = new W.SuppressLineNumbers();
                    if (_pPr != null)
                        _pPr.SuppressLineNumbers = _suppressLineNumbers;
                    else if (_spPr != null)
                        _spPr.SuppressLineNumbers = _suppressLineNumbers;
                }
                if (value)
                    _suppressLineNumbers.Val = null;
                else
                    _suppressLineNumbers.Val = false;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether suppress hyphenation for paragraph.
        /// </summary>
        public BooleanValue SuppressAutoHyphens
        {
            get
            {
                if (_suppressAutoHyphens == null) return null;
                if (_suppressAutoHyphens.Val == null) return true;
                return _suppressAutoHyphens.Val.Value;
            }
            set
            {
                if (_suppressAutoHyphens == null)
                {
                    _suppressAutoHyphens = new W.SuppressAutoHyphens();
                    if (_pPr != null)
                        _pPr.SuppressAutoHyphens = _suppressAutoHyphens;
                    else if (_spPr != null)
                        _spPr.SuppressAutoHyphens = _suppressAutoHyphens;
                }
                if (value)
                    _suppressAutoHyphens.Val = null;
                else
                    _suppressAutoHyphens.Val = false;
            }
        }
        #endregion

        #region Line Break
        /// <summary>
        /// Gets or sets a value indicating whether use asian rules for controlling first and last character.
        /// </summary>
        public BooleanValue Kinsoku
        {
            get
            {
                if (_kinsoku == null) return null;
                if (_kinsoku.Val == null) return true;
                return _kinsoku.Val.Value;
            }
            set
            {
                if (_kinsoku == null)
                {
                    _kinsoku = new W.Kinsoku();
                    if (_pPr != null)
                        _pPr.Kinsoku = _kinsoku;
                    else if (_spPr != null)
                        _spPr.Kinsoku = _kinsoku;
                }
                if (value)
                    _kinsoku.Val = null;
                else
                    _kinsoku.Val = false;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether allow latin text to wrap in the middle of a word.
        /// </summary>
        public BooleanValue WordWrap
        {
            get
            {
                if (_wordWrap == null) return null;
                if (_wordWrap.Val == null) return true;
                return _wordWrap.Val.Value;
            }
            set
            {
                if (_wordWrap == null)
                {
                    _wordWrap = new W.WordWrap();
                    if (_pPr != null)
                        _pPr.WordWrap = _wordWrap;
                    else if (_spPr != null)
                        _spPr.WordWrap = _wordWrap;
                }
                if (value)
                    _wordWrap.Val = null;
                else
                    _wordWrap.Val = false;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether allow hanging punctuation.
        /// </summary>
        public BooleanValue OverflowPunctuation
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
        #endregion

        #region Character Spacing
        /// <summary>
        /// Gets or sets a value indicating whether allow punctuation at the start of a line to compress.
        /// </summary>
        public BooleanValue TopLinePunctuation
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
        /// Gets or sets a value indicating whether automatically adjust space between Asian and Latin text.
        /// </summary>
        public BooleanValue AutoSpaceDE
        {
            get
            {
                if (_autoSpaceDE == null) return null;
                if (_autoSpaceDE.Val == null) return true;
                return _autoSpaceDE.Val.Value;
            }
            set
            {
                if (_autoSpaceDE == null)
                {
                    _autoSpaceDE = new W.AutoSpaceDE();
                    if (_pPr != null)
                        _pPr.AutoSpaceDE = _autoSpaceDE;
                    else if (_spPr != null)
                        _spPr.AutoSpaceDE = _autoSpaceDE;
                }
                if (value)
                    _autoSpaceDE.Val = null;
                else
                    _autoSpaceDE.Val = false;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether automatically adjust space between Asian text and numbers.
        /// </summary>
        public BooleanValue AutoSpaceDN
        {
            get
            {
                if (_autoSpaceDN == null) return null;
                if (_autoSpaceDN.Val == null) return true;
                return _autoSpaceDN.Val.Value;
            }
            set
            {
                if (_autoSpaceDN == null)
                {
                    _autoSpaceDN = new W.AutoSpaceDN();
                    if (_pPr != null)
                        _pPr.AutoSpaceDN = _autoSpaceDN;
                    else if (_spPr != null)
                        _spPr.AutoSpaceDN = _autoSpaceDN;
                }
                if (value)
                    _autoSpaceDN.Val = null;
                else
                    _autoSpaceDN.Val = false;
            }
        }
        #endregion

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
