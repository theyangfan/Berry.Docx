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
        private bool _beforeAutoSpacing = false;
        private bool _afterAutoSpacing = false;
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
        /// <summary>
        /// Gets or sets a value indicating whether spacing before is automatic.
        /// </summary>
        public bool BeforeAutoSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    // direct formatting
                    if (_directPHld.BeforeAutoSpacing != null) return _directPHld.BeforeAutoSpacing;
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.BeforeAutoSpacing != null) return inheritedStyle.BeforeAutoSpacing;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.BeforeAutoSpacing;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.BeforeAutoSpacing != null) return inheritedStyle.BeforeAutoSpacing;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.BeforeAutoSpacing;
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
        /// Gets or sets a value indicating whether spacing after is automatic.
        /// </summary>
        public bool AfterAutoSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    // direct formatting
                    if (_directPHld.AfterAutoSpacing != null) return _directPHld.AfterAutoSpacing;
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (inheritedStyle.AfterAutoSpacing != null) return inheritedStyle.AfterAutoSpacing;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.AfterAutoSpacing;
                }
                else if (_ownerStyle != null)
                {
                    // paragraph style inheritance
                    ParagraphPropertiesHolder inheritedStyle = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                    if (inheritedStyle.AfterAutoSpacing != null) return inheritedStyle.AfterAutoSpacing;
                    // document defaults
                    return _doc.DefaultFormat.ParagraphFormat.AfterAutoSpacing;
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
        /// <summary>
        /// Gets the left indent of the paragraph.
        /// </summary>
        /// <returns></returns>
        public Indentation GetLeftIndent()
        {
            Indentation ind = null;
            if (_ownerParagraph != null)
            {
                ind = ParagraphPropertiesHolder.GetParagraphLeftIndentation(_doc, _ownerParagraph);
            }
            else if(_ownerStyle != null)
            {
                ind = ParagraphPropertiesHolder.GetStyleLeftIndentation(_doc, _ownerStyle);
            }
            return ind ?? new Indentation(0, IndentationUnit.Character);
        }

        /// <summary>
        /// Sets the left indent for paragraph.
        /// </summary>
        /// <param name="val"></param>
        /// <param name="unit"></param>
        public void SetLeftIndent(float val, IndentationUnit unit)
        {
            if(_ownerParagraph != null)
            {
                SpecialIndentation hangingInd = ParagraphPropertiesHolder.GetParagraphSpecialPointsIndentation(_doc, _ownerParagraph);
                if(unit == IndentationUnit.Character)
                {
                    _directPHld.LeftIndent = val * 5;
                    if (hangingInd != null && hangingInd.Type == SpecialIndentationType.Hanging)
                    {
                        _directPHld.LeftIndent = val * 5 + hangingInd.Val;
                    }
                    _directPHld.LeftCharsIndent = val;
                }
                else
                {
                    _directPHld.LeftCharsIndent = 0;
                    if (hangingInd != null && hangingInd.Type == SpecialIndentationType.Hanging)
                    {
                        _directPHld.LeftIndent = val + hangingInd.Val;
                    }
                    else
                    {
                        _directPHld.LeftIndent = val;
                    }
                }
            }
            else if(_ownerStyle != null)
            {
                SpecialIndentation hangingInd = ParagraphPropertiesHolder.GetStyleSpecialPointsIndentation(_doc, _ownerStyle);
                if (unit == IndentationUnit.Character)
                {
                    _directSHld.LeftIndent = val * 5;
                    if (hangingInd != null && hangingInd.Type == SpecialIndentationType.Hanging)
                    {
                        _directSHld.LeftIndent = val * 5 + hangingInd.Val;
                    }
                    _directSHld.LeftCharsIndent = val;
                }
                else
                {
                    _directSHld.LeftCharsIndent = 0;
                    if (hangingInd != null && hangingInd.Type == SpecialIndentationType.Hanging)
                    {
                        _directSHld.LeftIndent = val + hangingInd.Val;
                    }
                    else
                    {
                        _directSHld.LeftIndent = val;
                    }
                }
            }
        }

        /// <summary>
        /// Gets the right indent of the paragraph.
        /// </summary>
        /// <returns></returns>
        public Indentation GetRightIndent()
        {
            Indentation ind = null;
            if (_ownerParagraph != null)
            {
                ind = ParagraphPropertiesHolder.GetParagraphRightIndentation(_doc, _ownerParagraph);
            }
            else if (_ownerStyle != null)
            {
                ind = ParagraphPropertiesHolder.GetStyleRightIndentation(_doc, _ownerStyle);
            }
            return ind ?? new Indentation(0, IndentationUnit.Character);
        }

        /// <summary>
        /// Sets the right indent for paragraph.
        /// </summary>
        /// <param name="val"></param>
        /// <param name="unit"></param>
        public void SetRightIndent(float val, IndentationUnit unit)
        {
            if (_ownerParagraph != null)
            {
                if (unit == IndentationUnit.Character)
                {
                    _directPHld.RightIndent = val * 5;
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
                    _directSHld.RightIndent = val * 5;
                    _directSHld.RightCharsIndent = val;
                }
                else
                {
                    _directSHld.RightCharsIndent = 0;
                    _directSHld.RightIndent = val;
                }
            }
        }

        /// <summary>
        /// Gets the special indent of the paragraph.
        /// </summary>
        /// <returns></returns>
        public SpecialIndentation GetSpecialIndentation()
        {
            SpecialIndentation ind = null;
            if (_ownerParagraph != null)
            {
                 ind = ParagraphPropertiesHolder.GetParagraphSpecialIndentation(_doc, _ownerParagraph);
            }
            else if(_ownerStyle != null)
            {
                ind = ParagraphPropertiesHolder.GetStyleSpecialIndentation(_doc, _ownerStyle);
            }
            return ind ?? new SpecialIndentation(SpecialIndentationType.None, 0, IndentationUnit.Character);
        }

        /// <summary>
        /// Sets the special indent for paragraph.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="val"></param>
        /// <param name="unit"></param>
        public void SetSpecialIndentation(SpecialIndentationType type, float val = 0, IndentationUnit unit = IndentationUnit.Character)
        {
            if(_ownerParagraph != null)
            {
                if (type == SpecialIndentationType.FirstLine)
                {
                    if(unit == IndentationUnit.Character)
                    {
                        _directPHld.HangingIndent = null;
                        _directPHld.HangingCharsIndent = null;
                        _directPHld.FirstLineIndent = val * 5;
                        _directPHld.FirstLineCharsIndent = val;
                    }
                    else
                    {
                        _directPHld.HangingIndent = null;
                        _directPHld.HangingCharsIndent = null;
                        _directPHld.FirstLineCharsIndent = 0;
                        _directPHld.FirstLineIndent = val;
                    }
                }
                else if(type == SpecialIndentationType.Hanging)
                {
                    if (unit == IndentationUnit.Character)
                    {
                        _directPHld.FirstLineIndent = null;
                        _directPHld.FirstLineCharsIndent = null;
                        _directPHld.HangingIndent = val * 5;
                        _directPHld.HangingCharsIndent = val;
                    }
                    else
                    {
                        Indentation ind = GetLeftIndent();
                        float leftTemp = ind.Unit == IndentationUnit.Point ? ind.Val : -1;

                        _directPHld.FirstLineIndent = null;
                        _directPHld.FirstLineCharsIndent = null;
                        _directPHld.HangingCharsIndent = 0;
                        _directPHld.HangingIndent = val;
                        // 重新设置左缩进
                        if(leftTemp >= 0) SetLeftIndent(leftTemp, IndentationUnit.Point);
                    }
                }
                else
                {
                    _directPHld.HangingIndent = null;
                    _directPHld.HangingCharsIndent = null;
                    _directPHld.FirstLineCharsIndent = 0;
                    _directPHld.FirstLineIndent = 0;
                }
            }
            else if(_ownerStyle != null)
            {
                if (type == SpecialIndentationType.FirstLine)
                {
                    if (unit == IndentationUnit.Character)
                    {
                        _directSHld.HangingIndent = null;
                        _directSHld.HangingCharsIndent = null;
                        _directSHld.FirstLineIndent = val * 5;
                        _directSHld.FirstLineCharsIndent = val;
                    }
                    else
                    {
                        _directSHld.HangingIndent = null;
                        _directSHld.HangingCharsIndent = null;
                        _directSHld.FirstLineCharsIndent = 0;
                        _directSHld.FirstLineIndent = val;
                    }
                }
                else if (type == SpecialIndentationType.Hanging)
                {
                    if (unit == IndentationUnit.Character)
                    {
                        _directSHld.FirstLineIndent = null;
                        _directSHld.FirstLineCharsIndent = null;
                        _directSHld.HangingIndent = val * 5;
                        _directSHld.HangingCharsIndent = val;
                    }
                    else
                    {
                        _directSHld.FirstLineIndent = null;
                        _directSHld.FirstLineCharsIndent = null;
                        _directSHld.HangingCharsIndent = 0;
                        _directSHld.HangingIndent = val;
                    }
                }
                else
                {
                    _directSHld.HangingIndent = null;
                    _directSHld.HangingCharsIndent = null;
                    _directSHld.FirstLineCharsIndent = 0;
                    _directSHld.FirstLineIndent = 0;
                }
            }
        }

        public Spacing GetBeforeSpacing()
        {
            FloatValue beforeSpacing = null;
            FloatValue beforeSpacingLines = null;
            if(_ownerParagraph != null)
            {
                ParagraphPropertiesHolder style = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                beforeSpacing = _directPHld.BeforeSpacing ?? style.BeforeSpacing;
                beforeSpacingLines = _directPHld.BeforeLinesSpacing ?? style.BeforeLinesSpacing;
            }
            else if(_ownerStyle != null)
            {
                ParagraphPropertiesHolder style = ParagraphPropertiesHolder.GetParagraphStyleFormatRecursively(_doc, _ownerStyle);
                beforeSpacing = style.BeforeSpacing;
                beforeSpacingLines = style.BeforeLinesSpacing;
            }
            if(beforeSpacingLines != null && beforeSpacingLines != 0)
            {
                return new Spacing(beforeSpacingLines, SpacingUnit.Line);
            }
            else if(beforeSpacing != null && beforeSpacing != 0)
            {
                return new Spacing(beforeSpacing, SpacingUnit.Point);
            }
            return new Spacing(0, SpacingUnit.Line);
        }

        public void SetBeforeSpacing(float val, SpacingUnit unit)
        {
            if(_ownerParagraph != null)
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
