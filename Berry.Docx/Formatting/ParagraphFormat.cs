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
        // Style Members
        private W.Style _ownerStyle;
        private ParagraphPropertiesHolder _directSHld;

        // Formats Members
        // Normal
        private JustificationType _justification = JustificationType.Both;
        private OutlineLevelType _outlineLevel = OutlineLevelType.BodyText;
        // Indentation
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.Justification != null) return inheritedStyle.Justification;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.Justification;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.Justification != null) return inheritedStyle.Justification;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.OutlineLevel != null) return inheritedStyle.OutlineLevel;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.OutlineLevel;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.OutlineLevel != null) return inheritedStyle.OutlineLevel;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.MirrorIndents != null) return inheritedStyle.MirrorIndents;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.MirrorIndents;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.MirrorIndents != null) return inheritedStyle.MirrorIndents;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.AdjustRightIndent != null) return inheritedStyle.AdjustRightIndent;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.AdjustRightIndent;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.AdjustRightIndent != null) return inheritedStyle.AdjustRightIndent;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.ContextualSpacing != null) return inheritedStyle.ContextualSpacing;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.ContextualSpacing;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.ContextualSpacing != null) return inheritedStyle.ContextualSpacing;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.SnapToGrid != null) return inheritedStyle.SnapToGrid;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.SnapToGrid;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.SnapToGrid != null) return inheritedStyle.SnapToGrid;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.WidowControl != null) return inheritedStyle.WidowControl;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.WidowControl;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.WidowControl != null) return inheritedStyle.WidowControl;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.KeepNext != null) return inheritedStyle.KeepNext;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.KeepNext;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.KeepNext != null) return inheritedStyle.KeepNext;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.KeepLines != null) return inheritedStyle.KeepLines;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.KeepLines;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.KeepLines != null) return inheritedStyle.KeepLines;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.PageBreakBefore != null) return inheritedStyle.PageBreakBefore;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.PageBreakBefore;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.PageBreakBefore != null) return inheritedStyle.PageBreakBefore;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.SuppressLineNumbers != null) return inheritedStyle.SuppressLineNumbers;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.SuppressLineNumbers;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.SuppressLineNumbers != null) return inheritedStyle.SuppressLineNumbers;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.SuppressAutoHyphens != null) return inheritedStyle.SuppressAutoHyphens;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.SuppressAutoHyphens;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.SuppressAutoHyphens != null) return inheritedStyle.SuppressAutoHyphens;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.Kinsoku != null) return inheritedStyle.Kinsoku;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.Kinsoku;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.Kinsoku != null) return inheritedStyle.Kinsoku;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.WordWrap != null) return inheritedStyle.WordWrap;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.WordWrap;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.WordWrap != null) return inheritedStyle.WordWrap;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.OverflowPunctuation != null) return inheritedStyle.OverflowPunctuation;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.OverflowPunctuation;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.OverflowPunctuation != null) return inheritedStyle.OverflowPunctuation;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.TopLinePunctuation != null) return inheritedStyle.TopLinePunctuation;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.TopLinePunctuation;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.TopLinePunctuation != null) return inheritedStyle.TopLinePunctuation;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.AutoSpaceDE != null) return inheritedStyle.AutoSpaceDE;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.AutoSpaceDE;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.AutoSpaceDE != null) return inheritedStyle.AutoSpaceDE;
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
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.AutoSpaceDN != null) return inheritedStyle.AutoSpaceDN;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.AutoSpaceDN;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.AutoSpaceDN != null) return inheritedStyle.AutoSpaceDN;
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
        public Indentation GetLeftIndent()
        {
            FloatValue leftInd = null;
            FloatValue leftIndChars = null;
            FloatValue hangingInd = null;
            FloatValue hangingIndChars = null;
            if (_ownerParagraph != null)
            {
                ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                leftInd = _directPHld.LeftIndent ?? inheritedStyle.LeftIndent;
                leftIndChars = _directPHld.LeftCharsIndent ?? inheritedStyle.LeftCharsIndent;
                hangingInd = _directPHld.HangingIndent?? inheritedStyle.HangingIndent;
                hangingIndChars = _directPHld.HangingCharsIndent ?? inheritedStyle.HangingCharsIndent;
            }
            else if (_ownerStyle != null)
            {
                ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                leftInd = inheritedStyle.LeftIndent;
                leftIndChars = inheritedStyle.LeftCharsIndent;
                hangingInd = inheritedStyle.HangingIndent;
                hangingIndChars = inheritedStyle.HangingCharsIndent;
            }
            // 字符
            if (leftIndChars != null && leftIndChars != 0)
            {
                return new Indentation(leftIndChars, IndentationUnit.Character);
            }
            // 磅
            if (hangingIndChars == null || hangingIndChars == 0)
            {
                if (hangingInd != null && hangingInd > 0)
                {
                    return new Indentation((leftInd ?? 0) - hangingInd, IndentationUnit.Point);
                }
                else
                {
                    if(leftInd != null && leftInd != 0) return new Indentation(leftInd, IndentationUnit.Point);
                }
            }
            return new Indentation(0, IndentationUnit.Character);
        }

        public void SetLeftIndent(float val, IndentationUnit unit)
        {
            if(_ownerParagraph != null)
            {
                ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                FloatValue hangingInd = _directPHld.HangingIndent ?? inheritedStyle.HangingIndent;
                FloatValue hangingIndChars = _directPHld.HangingCharsIndent ?? inheritedStyle.HangingCharsIndent;
                if(unit == IndentationUnit.Character)
                {
                    if (val == 0)
                    {
                        _directPHld.LeftIndent = 0;
                        if ((hangingIndChars == null || hangingIndChars == 0)
                            && (hangingInd != null && hangingInd > 0))
                            _directPHld.LeftIndent = hangingInd;
                    }
                    _directPHld.LeftCharsIndent = val;
                }
                else
                {
                    _directPHld.LeftCharsIndent = 0;
                    if ((hangingIndChars == null || hangingIndChars == 0)
                            && (hangingInd != null && hangingInd > 0))
                    {
                        _directPHld.LeftIndent = val + hangingInd;
                    }
                    else
                    {
                        _directPHld.LeftIndent = val;
                    }
                }
            }
            else if(_ownerStyle != null)
            {
                ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                FloatValue hangingInd = inheritedStyle.HangingIndent;
                FloatValue hangingIndChars = inheritedStyle.HangingCharsIndent;
                if (unit == IndentationUnit.Character)
                {
                    if (val == 0)
                    {
                        _directSHld.LeftIndent = 0;
                        if ((hangingIndChars == null || hangingIndChars == 0)
                            && (hangingInd != null && hangingInd > 0))
                            _directSHld.LeftIndent = hangingInd;
                    }
                    _directSHld.LeftCharsIndent = val;
                }
                else
                {
                    _directSHld.LeftCharsIndent = 0;
                    if ((hangingIndChars == null || hangingIndChars == 0)
                            && (hangingInd != null && hangingInd > 0))
                    {
                        _directSHld.LeftIndent = val + hangingInd;
                    }
                    else
                    {
                        _directSHld.LeftIndent = val;
                    }
                }
            }
        }

        public Indentation GetRightIndent()
        {
            FloatValue rightInd = null;
            FloatValue rightIndChars = null;
            if (_ownerParagraph != null)
            {
                ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                rightInd = _directPHld.RightIndent ?? inheritedStyle.RightIndent;
                rightIndChars = _directPHld.RightCharsIndent ?? inheritedStyle.RightCharsIndent;
            }
            else if (_ownerStyle != null)
            {
                ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                rightInd = inheritedStyle.RightIndent;
                rightIndChars = inheritedStyle.RightCharsIndent;
            }
            // 字符
            if (rightIndChars != null && rightIndChars != 0)
            {
                return new Indentation(rightIndChars, IndentationUnit.Character);
            }
            // 磅
            if (rightInd != null && rightInd != 0)
            {
                return new Indentation(rightInd, IndentationUnit.Point);
            }
            return new Indentation(0, IndentationUnit.Character);
        }

        public void SetRightIndent(float val, IndentationUnit unit)
        {
            if (_ownerParagraph != null)
            {
                if (unit == IndentationUnit.Character)
                {
                    if (val == 0)
                    {
                        _directPHld.RightIndent = 0;
                    }
                    _directPHld.RightCharsIndent = val;
                }
                else
                {
                    _directPHld.RightCharsIndent = 0;
                    _directPHld.RightIndent = val;
                }
            }
            else if (_ownerStyle != null)
            {
                if (unit == IndentationUnit.Character)
                {
                    if (val == 0)
                    {
                        _directSHld.RightIndent = 0;
                    }
                    _directSHld.RightCharsIndent = val;
                }
                else
                {
                    _directSHld.RightCharsIndent = 0;
                    _directSHld.RightIndent = val;
                }
            }
        }

        public SpecialIndentation GetSpecialIndentation()
        {
            FloatValue firstLineInd = null;
            FloatValue firstLineIndChars = null;
            FloatValue hangingInd = null;
            FloatValue hangingIndChars = null;
            if(_ownerParagraph != null)
            {
                ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                firstLineInd = _directPHld.FirstLineIndent ?? inheritedStyle.FirstLineIndent;
                firstLineIndChars = _directPHld.FirstLineCharsIndent ?? inheritedStyle.FirstLineCharsIndent;
                hangingInd = _directPHld.HangingIndent ?? inheritedStyle.HangingIndent;
                hangingIndChars = _directPHld.HangingCharsIndent ?? inheritedStyle.HangingCharsIndent;
            }
            else if(_ownerStyle != null)
            {
                ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                firstLineInd = inheritedStyle.FirstLineIndent;
                firstLineIndChars = inheritedStyle.FirstLineCharsIndent;
                hangingInd = inheritedStyle.HangingIndent;
                hangingIndChars = inheritedStyle.HangingCharsIndent;
            }
            if(hangingIndChars != null && hangingIndChars != 0)
            {
                if(hangingIndChars > 0)
                    return new SpecialIndentation(SpecialIndentationType.Hanging, hangingIndChars, IndentationUnit.Character);
                else
                    return new SpecialIndentation(SpecialIndentationType.FirstLine, -hangingIndChars, IndentationUnit.Character);
            } // hangingIndChars == null 或 0
            if (hangingInd != null && hangingInd != 0)
            {
                if((hangingIndChars != null && hangingIndChars == 0)
                    || firstLineIndChars == null || firstLineIndChars == 0)
                {
                    if (hangingInd > 0)
                        return new SpecialIndentation(SpecialIndentationType.Hanging, hangingInd, IndentationUnit.Point);
                    else
                        return new SpecialIndentation(SpecialIndentationType.FirstLine, -hangingInd, IndentationUnit.Point);
                }
            } // hangingInd == null 或 0
            if(hangingIndChars != null && hangingIndChars == 0 && hangingInd != null && hangingInd == 0)
            {
                return new SpecialIndentation(SpecialIndentationType.None, 0, IndentationUnit.Character);
            }
            if(firstLineIndChars != null && firstLineIndChars != 0 && hangingIndChars == null)
            {
                if (firstLineIndChars > 0)
                    return new SpecialIndentation(SpecialIndentationType.FirstLine, firstLineIndChars, IndentationUnit.Character);
                else
                    return new SpecialIndentation(SpecialIndentationType.Hanging, -firstLineIndChars, IndentationUnit.Character);
            }
            if(firstLineInd != null && firstLineInd != 0
                && hangingInd == null 
                || (hangingIndChars == null && (firstLineIndChars == null || firstLineIndChars == 0))
                || (hangingIndChars != null && hangingIndChars == 0))
            {
                if (firstLineInd > 0)
                    return new SpecialIndentation(SpecialIndentationType.FirstLine, firstLineInd, IndentationUnit.Point);
                else
                    return new SpecialIndentation(SpecialIndentationType.Hanging, -firstLineInd, IndentationUnit.Point);
            }
            return new SpecialIndentation(SpecialIndentationType.None, 0, IndentationUnit.Character);
        }

        public void SetSpecialIndentation(SpecialIndentationType type, float val, IndentationUnit unit)
        {
            if(_ownerParagraph != null)
            {
                if (type == SpecialIndentationType.FirstLine)
                {
                    if(unit == IndentationUnit.Character)
                    {
                        _directPHld.HangingIndent = 0;
                        _directPHld.HangingCharsIndent = 0;
                    }
                    else
                    {

                    }
                }
            }
            else if(_ownerStyle != null)
            {

            }
        }

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
