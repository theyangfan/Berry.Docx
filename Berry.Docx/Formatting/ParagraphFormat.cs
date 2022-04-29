using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OOxml = DocumentFormat.OpenXml.Wordprocessing;
using P = DocumentFormat.OpenXml.Packaging;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the paragraph format.
    /// </summary>
    public class ParagraphFormat
    {
        #region Private Members

        private Document _document = null;

        // Paragraph Members
        private OOxml.Paragraph _ownerParagraph = null;
        private ParagraphPropertiesHolder _curPHld = null;
        private ParagraphFormat _styleFormat = null;

        // Style Members
        private OOxml.Style _ownerStyle = null;
        private ParagraphPropertiesHolder _curSHld = null;
        private ParagraphFormat _styleHierarchyFormat = null;

        // Formats Members
        // Normal
        private JustificationType _justification = JustificationType.Both;
        private OutlineLevelType _outlineLevel = OutlineLevelType.BodyText;
        // Indentation
        private float _leftIndent = -1;
        private float _rightIndent = -1;
        private float _leftCharsIndent = -1;
        private float _rightCharsIndent = -1;
        private float _firstLineIndent = -1;
        private float _firstLineCharsIndent = -1;
        private float _hangingIndent = -1;
        private float _hangingCharsIndent = -1;
        private bool _mirrorIndents = false;
        private bool _adjustRightIndent = true;
        // Spacing
        private float _beforeSpacing = -1;
        private float _beforeLinesSpacing = -1;
        private bool _beforeAutoSpacing = false;
        private float _afterSpacing = -1;
        private float _afterLinesSpacing = -1;
        private bool _afterAutoSpacing = false;
        private float _lineSpacing = 12;
        private LineSpacingRule _lineSpacingRule = LineSpacingRule.Multiple;
        private bool _contextualSpacing = false;
        private bool _snapToGrid = true;
        // Pagination
        private bool _widowControl = false;
        private bool _keepNext = false;
        private bool _keepLines = false;
        private bool _pageBreakBefore = false;
        // Formatting Exceptions
        private bool _suppressLineNumbers = false;
        private bool _suppressAutoHyphens = false;
        // Line Break
        private bool _kinsoku = true;
        private bool _wordWrap = true;
        private bool _overflowPunctuation = true;
        // Character Spacing
        private bool _topLinePunctuation = false;
        private bool _autoSpaceDE = true;
        private bool _autoSpaceDN = true;
        // Numbering
        private NumberingFormat _numFormat = null;

        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the ParagraphFormat class. 
        /// </summary>
        internal ParagraphFormat() { }
        /// <summary>
        /// Represent the paragraph format of a Paragraph. 
        /// </summary>
        /// <param name="document"></param>
        /// <param name="ownerParagraph"></param>
        internal ParagraphFormat(Document document, OOxml.Paragraph ownerParagraph)
        {
            _document = document;
            _ownerParagraph = ownerParagraph;
            if (ownerParagraph.ParagraphProperties == null)
                ownerParagraph.ParagraphProperties = new OOxml.ParagraphProperties();
            _curPHld = new ParagraphPropertiesHolder(document, ownerParagraph.ParagraphProperties);
            _styleFormat = new ParagraphFormat(document, ownerParagraph.GetStyle(document));
        }

        /// <summary>
        /// Represent the paragraph format of a ParagraphStyle. 
        /// </summary>
        /// <param name="document"></param>
        /// <param name="ownerStyle"></param>
        internal ParagraphFormat(Document document, OOxml.Style ownerStyle)
        {
            _document = document;
            _ownerStyle = ownerStyle;
            _curSHld = new ParagraphPropertiesHolder(document, ownerStyle.StyleParagraphProperties);
            _styleHierarchyFormat = GetStyleParagraphFormatRecursively(ownerStyle);
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
                if(_ownerParagraph != null)
                {
                    return _ownerParagraph.ParagraphProperties?.NumberingProperties != null ? _curPHld.NumberingFormat : _styleFormat.NumberingFormat;
                }
                else if(_ownerStyle != null)
                {
                    return _styleHierarchyFormat.NumberingFormat;
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

        #region Normal
        /// <summary>
        /// Gets or sets the justification.
        /// </summary>
        public JustificationType Justification
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.Justification != JustificationType.None ? _curPHld.Justification : _styleFormat.Justification;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.Justification != JustificationType.None ? _curSHld.Justification : _styleHierarchyFormat.Justification;
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
        /// Gets or sets the outline level.
        /// </summary>
        public OutlineLevelType OutlineLevel
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.OutlineLevel != OutlineLevelType.None ? _curPHld.OutlineLevel : _styleFormat.OutlineLevel;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.OutlineLevel != OutlineLevelType.None ? _curSHld.OutlineLevel : _styleHierarchyFormat.OutlineLevel;
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
        #endregion

        #region Indentation
        /// <summary>
        /// Gets or sets the left indent (in points) for paragraph.
        /// </summary>
        public float LeftIndent
        {
            get
            {
                if(_ownerParagraph != null)
                {
                    return _curPHld.LeftIndent ?? _styleFormat.LeftIndent;
                }
                else if(_ownerStyle != null)
                {
                    return _curSHld.LeftIndent ?? _styleHierarchyFormat.LeftIndent;
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
        /// Gets or sets the left indent (in chars) for paragraph.
        /// </summary>
        public float LeftCharsIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.LeftCharsIndent ?? _styleFormat.LeftCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.LeftCharsIndent ?? _styleHierarchyFormat.LeftCharsIndent;
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
        /// Gets or sets the right indent (in points) for paragraph.
        /// </summary>
        public float RightIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.RightIndent ?? _styleFormat.RightIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.RightIndent ?? _styleHierarchyFormat.RightIndent;
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
        /// Gets or sets the right indent (in chars) for paragraph.
        /// </summary>
        public float RightCharsIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.RightCharsIndent ?? _styleFormat.RightCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.RightCharsIndent ?? _styleHierarchyFormat.RightCharsIndent;
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
        /// Gets or sets the first line indent (in points) for paragraph.
        /// </summary>
        public float FirstLineIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.FirstLineIndent ?? _styleFormat.FirstLineIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.FirstLineIndent ?? _styleHierarchyFormat.FirstLineIndent;
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
        /// Gets or sets the first line indent (in chars) for paragraph.
        /// </summary>
        public float FirstLineCharsIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.FirstLineCharsIndent ?? _styleFormat.FirstLineCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.FirstLineCharsIndent ?? _styleHierarchyFormat.FirstLineCharsIndent;
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
        /// Gets or sets the hanging indent (in points) for paragraph.
        /// </summary>
        public float HangingIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.HangingIndent ?? _styleFormat.HangingIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.HangingIndent ?? _styleHierarchyFormat.HangingIndent;
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
        /// Gets or sets the hanging indent (in chars) for paragraph.
        /// </summary>
        public float HangingCharsIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.HangingCharsIndent ?? _styleFormat.HangingCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.HangingCharsIndent ?? _styleHierarchyFormat.HangingCharsIndent;
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
        /// Gets or sets a value indicating whether automatically adjust right indent when document grid is defined.
        /// </summary>
        public bool AdjustRightIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.AdjustRightIndent ?? _styleFormat.AdjustRightIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.AdjustRightIndent ?? _styleHierarchyFormat.AdjustRightIndent;
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
        /// Gets or sets a value indicating whether the paragraph indents should be interpreted as mirrored indents.
        /// </summary>
        public bool MirrorIndents
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.MirrorIndents ?? _styleFormat.MirrorIndents;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.MirrorIndents ?? _styleHierarchyFormat.MirrorIndents;
                }
                else
                {
                    return _mirrorIndents;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.MirrorIndents = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.MirrorIndents = value;
                }
                else
                {
                    _mirrorIndents = value;
                }
            }
        }
        #endregion

        #region Spacing
        /// <summary>
        /// Gets or sets the spacing (in points) before the paragraph.
        /// </summary>
        public float BeforeSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.BeforeSpacing ?? _styleFormat.BeforeSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.BeforeSpacing ?? _styleHierarchyFormat.BeforeSpacing;
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
        /// Gets or sets the spacing (in lines) before the paragraph.
        /// </summary>
        public float BeforeLinesSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.BeforeLinesSpacing ?? _styleFormat.BeforeLinesSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.BeforeLinesSpacing ?? _styleHierarchyFormat.BeforeLinesSpacing;
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
        /// Gets or sets a value indicating whether spacing before is automatic.
        /// </summary>
        public bool BeforeAutoSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.BeforeAutoSpacing ?? _styleFormat.BeforeAutoSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.BeforeAutoSpacing ?? _styleHierarchyFormat.BeforeAutoSpacing;
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
        /// Gets or sets the spacing (in points) after the paragraph.
        /// </summary>
        public float AfterSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.AfterSpacing ??_styleFormat.AfterSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.AfterSpacing ?? _styleHierarchyFormat.AfterSpacing;
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
        /// Gets or sets the spacing (in lines) after the paragraph.
        /// </summary>
        public float AfterLinesSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.AfterLinesSpacing ?? _styleFormat.AfterLinesSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.AfterLinesSpacing ?? _styleHierarchyFormat.AfterLinesSpacing;
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
        /// Gets or sets a value indicating whether spacing after is automatic.
        /// </summary>
        public bool AfterAutoSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.AfterAutoSpacing ?? _styleFormat.AfterAutoSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.AfterAutoSpacing ?? _styleHierarchyFormat.AfterAutoSpacing;
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
        /// Gets or sets the line spacing (in points) for paragraph.
        /// </summary>
        public float LineSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.LineSpacing ?? _styleFormat.LineSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.LineSpacing ?? _styleHierarchyFormat.LineSpacing;
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
        /// Gets or sets the line spacing rule of paragraph.
        /// </summary>
        public LineSpacingRule LineSpacingRule
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.LineSpacingRule != LineSpacingRule.None ? _curPHld.LineSpacingRule : _styleFormat.LineSpacingRule;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.LineSpacingRule != LineSpacingRule.None ? _curSHld.LineSpacingRule : _styleHierarchyFormat.LineSpacingRule;
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
        /// Gets or sets a value indicating whether don't add space between paragraphs of the same style.
        /// </summary>
        public bool ContextualSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.ContextualSpacing ?? _styleFormat.ContextualSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.ContextualSpacing ?? _styleHierarchyFormat.ContextualSpacing;
                }
                else
                {
                    return _contextualSpacing;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.ContextualSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.ContextualSpacing = value;
                }
                else
                {
                    _contextualSpacing = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether snap to grid when document grid is defined.
        /// </summary>
        public bool SnapToGrid
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.SnapToGrid ?? _styleFormat.SnapToGrid;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.SnapToGrid ?? _styleHierarchyFormat.SnapToGrid;
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
        #endregion

        #region Pagination
        /// <summary>
        /// Gets or sets a value indicating whether a consumer shall prevent first/last line of this paragraph 
        /// from being displayed on a separate page.
        /// </summary>
        public bool WidowControl
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.WidowControl ?? _styleFormat.WidowControl;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.WidowControl ?? _styleHierarchyFormat.WidowControl;
                }
                else
                {
                    return _widowControl;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.WidowControl = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.WidowControl = value;
                }
                else
                {
                    _widowControl = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether keep paragraph with next paragraph on the same page.
        /// </summary>
        public bool KeepNext
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.KeepNext ?? _styleFormat.KeepNext;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.KeepNext ?? _styleHierarchyFormat.KeepNext;
                }
                else
                {
                    return _keepNext;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.KeepNext = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.KeepNext = value;
                }
                else
                {
                    _keepNext = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether keep all lines of this paragraph on one page.
        /// </summary>
        public bool KeepLines
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.KeepLines ?? _styleFormat.KeepLines;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.KeepLines ?? _styleHierarchyFormat.KeepLines;
                }
                else
                {
                    return _keepLines;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.KeepLines = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.KeepLines = value;
                }
                else
                {
                    _keepLines = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether start paragraph on next page.
        /// </summary>
        public bool PageBreakBefore
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.PageBreakBefore ?? _styleFormat.PageBreakBefore;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.PageBreakBefore ?? _styleHierarchyFormat.PageBreakBefore;
                }
                else
                {
                    return _pageBreakBefore;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.PageBreakBefore = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.PageBreakBefore = value;
                }
                else
                {
                    _pageBreakBefore = value;
                }
            }
        }
        #endregion

        #region Formatting Exceptions
        /// <summary>
        /// Gets or sets a value indicating whether suppress line numbers for paragraph.
        /// </summary>
        public bool SuppressLineNumbers
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.SuppressLineNumbers ?? _styleFormat.SuppressLineNumbers;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.SuppressLineNumbers ?? _styleHierarchyFormat.SuppressLineNumbers;
                }
                else
                {
                    return _suppressLineNumbers;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.SuppressLineNumbers = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.SuppressLineNumbers = value;
                }
                else
                {
                    _suppressLineNumbers = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether suppress hyphenation for paragraph.
        /// </summary>
        public bool SuppressAutoHyphens
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.SuppressAutoHyphens ?? _styleFormat.SuppressAutoHyphens;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.SuppressAutoHyphens ?? _styleHierarchyFormat.SuppressAutoHyphens;
                }
                else
                {
                    return _suppressAutoHyphens;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.SuppressAutoHyphens = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.SuppressAutoHyphens = value;
                }
                else
                {
                    _suppressAutoHyphens = value;
                }
            }
        }
        #endregion

        #region Line Break
        /// <summary>
        /// Gets or sets a value indicating whether use asian rules for controlling first and last character.
        /// </summary>
        public bool Kinsoku
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.Kinsoku ?? _styleFormat.Kinsoku;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.Kinsoku ?? _styleHierarchyFormat.Kinsoku;
                }
                else
                {
                    return _kinsoku;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.Kinsoku = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.Kinsoku = value;
                }
                else
                {
                    _kinsoku = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether allow latin text to wrap in the middle of a word.
        /// </summary>
        public bool WordWrap
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.WordWrap ?? _styleFormat.WordWrap;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.WordWrap ?? _styleHierarchyFormat.WordWrap;
                }
                else
                {
                    return _wordWrap;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.WordWrap = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.WordWrap = value;
                }
                else
                {
                    _wordWrap = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether allow hanging punctuation.
        /// </summary>
        public bool OverflowPunctuation
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.OverflowPunctuation ?? _styleFormat.OverflowPunctuation;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.OverflowPunctuation ?? _styleHierarchyFormat.OverflowPunctuation;
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
        #endregion

        #region Character Spacing
        /// <summary>
        /// Gets or sets a value indicating whether allow punctuation at the start of a line to compress.
        /// </summary>
        public bool TopLinePunctuation
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.TopLinePunctuation ?? _styleFormat.TopLinePunctuation;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.TopLinePunctuation ?? _styleHierarchyFormat.TopLinePunctuation;
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
        /// Gets or sets a value indicating whether automatically adjust space between Asian and Latin text.
        /// </summary>
        public bool AutoSpaceDE
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.AutoSpaceDE ?? _styleFormat.AutoSpaceDE;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.AutoSpaceDE ?? _styleHierarchyFormat.AutoSpaceDE;
                }
                else
                {
                    return _autoSpaceDE;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.AutoSpaceDE = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.AutoSpaceDE = value;
                }
                else
                {
                    _autoSpaceDE = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether automatically adjust space between Asian text and numbers.
        /// </summary>
        public bool AutoSpaceDN
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.AutoSpaceDN ?? _styleFormat.AutoSpaceDN;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.AutoSpaceDN ?? _styleHierarchyFormat.AutoSpaceDN;
                }
                else
                {
                    return _autoSpaceDN;
                }
            }
            set
            {
                if (_ownerParagraph != null)
                {
                    _curPHld.AutoSpaceDN = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.AutoSpaceDN = value;
                }
                else
                {
                    _autoSpaceDN = value;
                }
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
            if (_ownerParagraph != null)
            {
                _curPHld.RemoveFrame();
            }
            else if (_ownerStyle != null)
            {
                _curSHld.RemoveFrame();
            }
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Returns the paragraph format that specified in the style hierarchy of a style.
        /// </summary>
        /// <param name="style"> The style</param>
        /// <returns>The paragraph format that specified in the style hierarchy.</returns>
        private ParagraphFormat GetStyleParagraphFormatRecursively(OOxml.Style style)
        {
            ParagraphFormat format = new ParagraphFormat();
            ParagraphFormat baseFormat = new ParagraphFormat();
            // Gets base style format.
            OOxml.Style baseStyle = style.GetBaseStyle(_document);
            if (baseStyle != null)
                baseFormat = GetStyleParagraphFormatRecursively(baseStyle);

            ParagraphPropertiesHolder curSHld = new ParagraphPropertiesHolder(_document, style.StyleParagraphProperties);
            // Normal
            format.Justification = curSHld.Justification != JustificationType.None ? curSHld.Justification : baseFormat.Justification;
            format.OutlineLevel = curSHld.OutlineLevel != OutlineLevelType.None ? curSHld.OutlineLevel : baseFormat.OutlineLevel;
            // Indentation
            format.LeftIndent = curSHld.LeftIndent ?? baseFormat.LeftIndent;
            format.LeftCharsIndent = curSHld.LeftCharsIndent ?? baseFormat.LeftCharsIndent;
            format.RightIndent = curSHld.RightIndent ?? baseFormat.RightIndent;
            format.RightCharsIndent = curSHld.RightCharsIndent ?? baseFormat.RightCharsIndent;
            format.FirstLineIndent = curSHld.FirstLineIndent ?? baseFormat.FirstLineIndent;
            format.FirstLineCharsIndent = curSHld.FirstLineCharsIndent ?? baseFormat.FirstLineCharsIndent;
            format.HangingIndent = curSHld.HangingIndent ?? baseFormat.HangingIndent;
            format.HangingCharsIndent = curSHld.HangingCharsIndent ?? baseFormat.HangingCharsIndent;
            format.MirrorIndents = curSHld.MirrorIndents ?? baseFormat.MirrorIndents;
            format.AdjustRightIndent = curSHld.AdjustRightIndent ?? baseFormat.AdjustRightIndent;
            // Spacing
            format.BeforeSpacing = curSHld.BeforeSpacing ?? baseFormat.BeforeSpacing;
            format.BeforeLinesSpacing = curSHld.BeforeLinesSpacing ?? baseFormat.BeforeLinesSpacing;
            format.BeforeAutoSpacing = curSHld.BeforeAutoSpacing ?? baseFormat.BeforeAutoSpacing;
            format.AfterSpacing = curSHld.AfterSpacing ?? baseFormat.AfterSpacing;
            format.AfterLinesSpacing = curSHld.AfterLinesSpacing ?? baseFormat.AfterLinesSpacing;
            format.AfterAutoSpacing = curSHld.AfterAutoSpacing ?? baseFormat.AfterAutoSpacing;
            format.LineSpacing = curSHld.LineSpacing ?? baseFormat.LineSpacing;
            format.LineSpacingRule = curSHld.LineSpacingRule != LineSpacingRule.None
                ? curSHld.LineSpacingRule : baseFormat.LineSpacingRule;
            format.ContextualSpacing = curSHld.ContextualSpacing ?? baseFormat.ContextualSpacing;
            format.SnapToGrid = curSHld.SnapToGrid ?? baseFormat.SnapToGrid;
            // Pagination
            format.WidowControl = curSHld.WidowControl ?? baseFormat.WidowControl;
            format.KeepNext = curSHld.KeepNext ?? baseFormat.KeepNext;
            format.KeepLines = curSHld.KeepLines ?? baseFormat.KeepLines;
            format.PageBreakBefore = curSHld.PageBreakBefore ?? baseFormat.PageBreakBefore;
            // Formatting Exceptions
            format.SuppressLineNumbers = curSHld.SuppressLineNumbers ?? baseFormat.SuppressLineNumbers;
            format.SuppressAutoHyphens = curSHld.SuppressAutoHyphens ?? baseFormat.SuppressAutoHyphens;
            // Line Break
            format.Kinsoku = curSHld.Kinsoku ?? baseFormat.Kinsoku;
            format.WordWrap = curSHld.WordWrap ?? baseFormat.WordWrap;
            format.OverflowPunctuation = curSHld.OverflowPunctuation ?? baseFormat.OverflowPunctuation;
            // Character Spacing
            format.TopLinePunctuation = curSHld.TopLinePunctuation ?? baseFormat.TopLinePunctuation;
            format.AutoSpaceDE = curSHld.AutoSpaceDE ?? baseFormat.AutoSpaceDE;
            format.AutoSpaceDN = curSHld.AutoSpaceDN ?? baseFormat.AutoSpaceDN;
            // Numbering
            if(curSHld.NumberingFormat != null)
            {
                format.NumberingFormat = curSHld.NumberingFormat;
            }
            else if(baseFormat.NumberingFormat != null)
            {
                format.NumberingFormat = new NumberingFormat(_document, baseFormat.NumberingFormat, style.StyleId);
            }
            
            return format;
        }
        #endregion
    }
}
