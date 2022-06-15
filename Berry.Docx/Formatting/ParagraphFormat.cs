using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;
using P = DocumentFormat.OpenXml.Packaging;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the paragraph format.
    /// </summary>
    public class ParagraphFormat
    {
        #region Private Members

        private Document _doc = null;

        // Paragraph Members
        private W.Paragraph _ownerParagraph;
        private ParagraphPropertiesHolder _directPHld;
        private ParagraphPropertiesHolder _pStyleHld;

        // Style Members
        private W.Style _ownerStyle;
        private ParagraphPropertiesHolder _directSHld;

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
        internal ParagraphFormat(Document document, W.Paragraph ownerParagraph)
        {
            _doc = document;
            _ownerParagraph = ownerParagraph;
            _directPHld = new ParagraphPropertiesHolder(document, ownerParagraph);
            _pStyleHld = new ParagraphPropertiesHolder(document, ownerParagraph.GetStyle(document));
        }

        /// <summary>
        /// Represent the paragraph format of a ParagraphStyle. 
        /// </summary>
        /// <param name="document"></param>
        /// <param name="ownerStyle"></param>
        internal ParagraphFormat(Document document, W.Style ownerStyle)
        {
            _doc = document;
            _ownerStyle = ownerStyle;
            _directSHld = new ParagraphPropertiesHolder(document, ownerStyle);
        }

        #endregion

        #region Public Properties
        /// <summary>
        /// Gets paragraph numbering format.
        /// </summary>
        /*public NumberingFormat NumberingFormat
        {
            get
            {
                if(_ownerParagraph != null)
                {
                    return _ownerParagraph.ParagraphProperties?.NumberingProperties != null ? _directPHld.NumberingFormat : _styleFormat.NumberingFormat;
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
*/
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
                    // direct formatting
                    if (_directPHld.Justification != null) return _directPHld.Justification;
                    // paragraph style
                    if (_pStyleHld.Justification != null) return _pStyleHld.Justification;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.Justification;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.Justification != null) return _directSHld.Justification;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.Justification;
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
                    _directPHld.Justification = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.Justification = value;
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
                    // direct formatting
                    if (_directPHld.OutlineLevel != null) return _directPHld.OutlineLevel;
                    // paragraph style
                    if (_pStyleHld.OutlineLevel != null) return _pStyleHld.OutlineLevel;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.OutlineLevel;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.OutlineLevel != null) return _directSHld.OutlineLevel;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.OutlineLevel;
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
                    _directPHld.OutlineLevel = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.OutlineLevel = value;
                }
                else
                {
                    _outlineLevel = value;
                }
            }
        }
        #endregion

        #region Indentation
        /*
        /// <summary>
        /// Gets or sets the left indent (in points) for paragraph.
        /// </summary>
        public float LeftIndent
        {
            get
            {
                if(_ownerParagraph != null)
                {
                    return _directPHld.LeftIndent ?? _styleFormat.LeftIndent;
                }
                else if(_ownerStyle != null)
                {
                    return _directSHld.LeftIndent ?? _styleHierarchyFormat.LeftIndent;
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
                    _directPHld.LeftIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.LeftIndent = value;
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
                    return _directPHld.LeftCharsIndent ?? _styleFormat.LeftCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.LeftCharsIndent ?? _styleHierarchyFormat.LeftCharsIndent;
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
                    _directPHld.LeftCharsIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.LeftCharsIndent = value;
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
                    return _directPHld.RightIndent ?? _styleFormat.RightIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.RightIndent ?? _styleHierarchyFormat.RightIndent;
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
                    _directPHld.RightIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.RightIndent = value;
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
                    return _directPHld.RightCharsIndent ?? _styleFormat.RightCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.RightCharsIndent ?? _styleHierarchyFormat.RightCharsIndent;
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
                    _directPHld.RightCharsIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.RightCharsIndent = value;
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
                    return _directPHld.FirstLineIndent ?? _styleFormat.FirstLineIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.FirstLineIndent ?? _styleHierarchyFormat.FirstLineIndent;
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
                    _directPHld.FirstLineIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.FirstLineIndent = value;
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
                    return _directPHld.FirstLineCharsIndent ?? _styleFormat.FirstLineCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.FirstLineCharsIndent ?? _styleHierarchyFormat.FirstLineCharsIndent;
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
                    _directPHld.FirstLineCharsIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.FirstLineCharsIndent = value;
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
                    return _directPHld.HangingIndent ?? _styleFormat.HangingIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.HangingIndent ?? _styleHierarchyFormat.HangingIndent;
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
                    _directPHld.HangingIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.HangingIndent = value;
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
                    return _directPHld.HangingCharsIndent ?? _styleFormat.HangingCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.HangingCharsIndent ?? _styleHierarchyFormat.HangingCharsIndent;
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
                    _directPHld.HangingCharsIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.HangingCharsIndent = value;
                }
                else
                {
                    _hangingCharsIndent = value;
                }
            }
        }
        */

        /// <summary>
        /// Gets or sets a value indicating whether the paragraph indents should be interpreted as mirrored indents.
        /// </summary>
        public bool MirrorIndents
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    // direct formatting
                    if (_directPHld.MirrorIndents != null) return _directPHld.MirrorIndents;
                    // paragraph style
                    if (_pStyleHld.MirrorIndents != null) return _pStyleHld.MirrorIndents;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.MirrorIndents;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.MirrorIndents != null) return _directSHld.MirrorIndents;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.MirrorIndents;
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
                    _directPHld.MirrorIndents = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.MirrorIndents = value;
                }
                else
                {
                    _mirrorIndents = value;
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
                    // direct formatting
                    if (_directPHld.AdjustRightIndent != null) return _directPHld.AdjustRightIndent;
                    // paragraph style
                    if (_pStyleHld.AdjustRightIndent != null) return _pStyleHld.AdjustRightIndent;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.AdjustRightIndent;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.AdjustRightIndent != null) return _directSHld.AdjustRightIndent;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.AdjustRightIndent;
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
                    _directPHld.AdjustRightIndent = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.AdjustRightIndent = value;
                }
                else
                {
                    _adjustRightIndent = value;
                }
            }
        }
        #endregion

        #region Spacing
        /*
        /// <summary>
        /// Gets or sets the spacing (in points) before the paragraph.
        /// </summary>
        public float BeforeSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _directPHld.BeforeSpacing ?? _styleFormat.BeforeSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.BeforeSpacing ?? _styleHierarchyFormat.BeforeSpacing;
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
                    _directPHld.BeforeSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.BeforeSpacing = value;
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
                    return _directPHld.BeforeLinesSpacing ?? _styleFormat.BeforeLinesSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.BeforeLinesSpacing ?? _styleHierarchyFormat.BeforeLinesSpacing;
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
                    _directPHld.BeforeLinesSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.BeforeLinesSpacing = value;
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
                    return _directPHld.BeforeAutoSpacing ?? _styleFormat.BeforeAutoSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.BeforeAutoSpacing ?? _styleHierarchyFormat.BeforeAutoSpacing;
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
                    _directPHld.BeforeAutoSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.BeforeAutoSpacing = value;
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
                    return _directPHld.AfterSpacing ??_styleFormat.AfterSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.AfterSpacing ?? _styleHierarchyFormat.AfterSpacing;
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
                    _directPHld.AfterSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.AfterSpacing = value;
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
                    return _directPHld.AfterLinesSpacing ?? _styleFormat.AfterLinesSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.AfterLinesSpacing ?? _styleHierarchyFormat.AfterLinesSpacing;
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
                    _directPHld.AfterLinesSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.AfterLinesSpacing = value;
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
                    return _directPHld.AfterAutoSpacing ?? _styleFormat.AfterAutoSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.AfterAutoSpacing ?? _styleHierarchyFormat.AfterAutoSpacing;
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
                    _directPHld.AfterAutoSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.AfterAutoSpacing = value;
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
                    return _directPHld.LineSpacing ?? _styleFormat.LineSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.LineSpacing ?? _styleHierarchyFormat.LineSpacing;
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
                    _directPHld.LineSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.LineSpacing = value;
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
                    return _directPHld.LineSpacingRule != LineSpacingRule.None ? _directPHld.LineSpacingRule : _styleFormat.LineSpacingRule;
                }
                else if (_ownerStyle != null)
                {
                    return _directSHld.LineSpacingRule != LineSpacingRule.None ? _directSHld.LineSpacingRule : _styleHierarchyFormat.LineSpacingRule;
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
                    _directPHld.LineSpacingRule = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.LineSpacingRule = value;
                }
                else
                {
                    _lineSpacingRule = value;
                }
            }
        }
        */
        /// <summary>
        /// Gets or sets a value indicating whether don't add space between paragraphs of the same style.
        /// </summary>
        public bool ContextualSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    // direct formatting
                    if (_directPHld.ContextualSpacing != null) return _directPHld.ContextualSpacing;
                    // paragraph style
                    if (_pStyleHld.ContextualSpacing != null) return _pStyleHld.ContextualSpacing;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.ContextualSpacing;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.ContextualSpacing != null) return _directSHld.ContextualSpacing;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.ContextualSpacing;
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
                    _directPHld.ContextualSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.ContextualSpacing = value;
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
                    // direct formatting
                    if (_directPHld.SnapToGrid != null) return _directPHld.SnapToGrid;
                    // paragraph style
                    if (_pStyleHld.SnapToGrid != null) return _pStyleHld.SnapToGrid;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.SnapToGrid;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.SnapToGrid != null) return _directSHld.SnapToGrid;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.SnapToGrid;
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
                    _directPHld.SnapToGrid = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.SnapToGrid = value;
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
                    // direct formatting
                    if (_directPHld.WidowControl != null) return _directPHld.WidowControl;
                    // paragraph style
                    if (_pStyleHld.WidowControl != null) return _pStyleHld.WidowControl;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.WidowControl;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.WidowControl != null) return _directSHld.WidowControl;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.WidowControl;
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
                    _directPHld.WidowControl = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.WidowControl = value;
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
                    // direct formatting
                    if (_directPHld.KeepNext != null) return _directPHld.KeepNext;
                    // paragraph style
                    if (_pStyleHld.KeepNext != null) return _pStyleHld.KeepNext;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.KeepNext;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.KeepNext != null) return _directSHld.KeepNext;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.KeepNext;
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
                    _directPHld.KeepNext = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.KeepNext = value;
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
                    // direct formatting
                    if (_directPHld.KeepLines != null) return _directPHld.KeepLines;
                    // paragraph style
                    if (_pStyleHld.KeepLines != null) return _pStyleHld.KeepLines;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.KeepLines;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.KeepLines != null) return _directSHld.KeepLines;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.KeepLines;
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
                    _directPHld.KeepLines = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.KeepLines = value;
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
                    // direct formatting
                    if (_directPHld.PageBreakBefore != null) return _directPHld.PageBreakBefore;
                    // paragraph style
                    if (_pStyleHld.PageBreakBefore != null) return _pStyleHld.PageBreakBefore;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.PageBreakBefore;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.PageBreakBefore != null) return _directSHld.PageBreakBefore;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.PageBreakBefore;
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
                    _directPHld.PageBreakBefore = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.PageBreakBefore = value;
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
                    // direct formatting
                    if (_directPHld.SuppressLineNumbers != null) return _directPHld.SuppressLineNumbers;
                    // paragraph style
                    if (_pStyleHld.SuppressLineNumbers != null) return _pStyleHld.SuppressLineNumbers;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.SuppressLineNumbers;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.SuppressLineNumbers != null) return _directSHld.SuppressLineNumbers;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.SuppressLineNumbers;
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
                    _directPHld.SuppressLineNumbers = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.SuppressLineNumbers = value;
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
                    // direct formatting
                    if (_directPHld.SuppressAutoHyphens != null) return _directPHld.SuppressAutoHyphens;
                    // paragraph style
                    if (_pStyleHld.SuppressAutoHyphens != null) return _pStyleHld.SuppressAutoHyphens;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.SuppressAutoHyphens;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.SuppressAutoHyphens != null) return _directSHld.SuppressAutoHyphens;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.SuppressAutoHyphens;
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
                    _directPHld.SuppressAutoHyphens = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.SuppressAutoHyphens = value;
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
                    // direct formatting
                    if (_directPHld.Kinsoku != null) return _directPHld.Kinsoku;
                    // paragraph style
                    if (_pStyleHld.Kinsoku != null) return _pStyleHld.Kinsoku;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.Kinsoku;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.Kinsoku != null) return _directSHld.Kinsoku;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.Kinsoku;
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
                    _directPHld.Kinsoku = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.Kinsoku = value;
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
                    // direct formatting
                    if (_directPHld.WordWrap != null) return _directPHld.WordWrap;
                    // paragraph style
                    if (_pStyleHld.WordWrap != null) return _pStyleHld.WordWrap;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.WordWrap;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.WordWrap != null) return _directSHld.WordWrap;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.WordWrap;
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
                    _directPHld.WordWrap = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.WordWrap = value;
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
                    // direct formatting
                    if (_directPHld.OverflowPunctuation != null) return _directPHld.OverflowPunctuation;
                    // paragraph style
                    if (_pStyleHld.OverflowPunctuation != null) return _pStyleHld.OverflowPunctuation;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.OverflowPunctuation;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.OverflowPunctuation != null) return _directSHld.OverflowPunctuation;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.OverflowPunctuation;
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
                    _directPHld.OverflowPunctuation = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.OverflowPunctuation = value;
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
                    // direct formatting
                    if (_directPHld.TopLinePunctuation != null) return _directPHld.TopLinePunctuation;
                    // paragraph style
                    if (_pStyleHld.TopLinePunctuation != null) return _pStyleHld.TopLinePunctuation;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.TopLinePunctuation;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.TopLinePunctuation != null) return _directSHld.TopLinePunctuation;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.TopLinePunctuation;
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
                    _directPHld.TopLinePunctuation = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.TopLinePunctuation = value;
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
                    // direct formatting
                    if (_directPHld.AutoSpaceDE != null) return _directPHld.AutoSpaceDE;
                    // paragraph style
                    if (_pStyleHld.AutoSpaceDE != null) return _pStyleHld.AutoSpaceDE;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.AutoSpaceDE;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.AutoSpaceDE != null) return _directSHld.AutoSpaceDE;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.AutoSpaceDE;
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
                    _directPHld.AutoSpaceDE = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.AutoSpaceDE = value;
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
                    // direct formatting
                    if (_directPHld.AutoSpaceDN != null) return _directPHld.AutoSpaceDN;
                    // paragraph style
                    if (_pStyleHld.AutoSpaceDN != null) return _pStyleHld.AutoSpaceDN;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.AutoSpaceDN;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style
                    if (_directSHld.AutoSpaceDN != null) return _directSHld.AutoSpaceDN;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.AutoSpaceDN;
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
                    _directPHld.AutoSpaceDN = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.AutoSpaceDN = value;
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
                _directPHld.RemoveFrame();
            }
            else if (_ownerStyle != null)
            {
                _directSHld.RemoveFrame();
            }
        }
        #endregion

    }
}
