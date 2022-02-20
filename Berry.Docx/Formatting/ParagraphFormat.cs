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

        #region Private Paragraph Members
        private OOxml.Paragraph _ownerParagraph = null;
        private ParagraphPropertiesHolder _curPHld = null;
        private ParagraphFormat _inheritFromStyleFormat = null;
        #endregion

        #region Private Style Members
        private OOxml.Style _ownerStyle = null;
        private ParagraphPropertiesHolder _curSHld = null;
        private ParagraphFormat _inheritFromBaseStyleFormat = null;
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

        private ZBool _overflowPunctuation = true;
        private ZBool _topLinePunctuation = false;
        private ZBool _adjustRightIndent = true;
        private ZBool _snapToGrid = true;
        #endregion

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
            _inheritFromStyleFormat = new ParagraphFormat(document, ownerParagraph.GetStyle(document));
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
            _inheritFromBaseStyleFormat = GetStyleParagraphFormatRecursively(ownerStyle);
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
                    return _ownerParagraph.ParagraphProperties != null && _ownerParagraph.ParagraphProperties.NumberingProperties != null ? _curPHld.NumberingFormat : _inheritFromStyleFormat.NumberingFormat;
                }
                else if(_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.NumberingFormat;
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
        /// Gets or sets the justification.
        /// </summary>
        public JustificationType Justification
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.Justification != JustificationType.None ? _curPHld.Justification : _inheritFromStyleFormat.Justification;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.Justification;
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
                    return _curPHld.OutlineLevel != OutlineLevelType.None ? _curPHld.OutlineLevel : _inheritFromStyleFormat.OutlineLevel;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.OutlineLevel;
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
        /// Gets or sets the left indent (in points) for paragraph.
        /// </summary>
        public float LeftIndent
        {
            get
            {
                if(_ownerParagraph != null)
                {
                    return _curPHld.LeftIndent >= 0 ? _curPHld.LeftIndent : _inheritFromStyleFormat.LeftIndent;
                }
                else if(_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.LeftIndent;
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
                    return _curPHld.LeftCharsIndent >= 0 ? _curPHld.LeftCharsIndent : _inheritFromStyleFormat.LeftCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.LeftCharsIndent;
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
                    return _curPHld.RightIndent >= 0 ? _curPHld.RightIndent : _inheritFromStyleFormat.RightIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.RightIndent;
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
                    return _curPHld.RightCharsIndent >= 0 ? _curPHld.RightCharsIndent : _inheritFromStyleFormat.RightCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.RightCharsIndent;
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
                    return _curPHld.FirstLineIndent >= 0 ? _curPHld.FirstLineIndent : _inheritFromStyleFormat.FirstLineIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.FirstLineIndent;
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
                    return _curPHld.FirstLineCharsIndent >= 0 ? _curPHld.FirstLineCharsIndent : _inheritFromStyleFormat.FirstLineCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.FirstLineCharsIndent;
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
                    return _curPHld.HangingIndent >= 0 ? _curPHld.HangingIndent : _inheritFromStyleFormat.HangingIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.HangingIndent;
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
                    return _curPHld.HangingCharsIndent >= 0 ? _curPHld.HangingCharsIndent : _inheritFromStyleFormat.HangingCharsIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.HangingCharsIndent;
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
        /// Gets or sets the spacing (in points) before the paragraph.
        /// </summary>
        public float BeforeSpacing
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.BeforeSpacing >= 0 ? _curPHld.BeforeSpacing : _inheritFromStyleFormat.BeforeSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.BeforeSpacing;
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
                    return _curPHld.BeforeLinesSpacing >= 0 ? _curPHld.BeforeLinesSpacing : _inheritFromStyleFormat.BeforeLinesSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.BeforeLinesSpacing;
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
                    return _curPHld.BeforeAutoSpacing ?? _inheritFromStyleFormat.BeforeAutoSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.BeforeAutoSpacing;
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
                    return _curPHld.AfterSpacing >= 0 ? _curPHld.AfterSpacing : _inheritFromStyleFormat.AfterSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.AfterSpacing;
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
                    return _curPHld.AfterLinesSpacing >= 0 ? _curPHld.AfterLinesSpacing : _inheritFromStyleFormat.AfterLinesSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.AfterLinesSpacing;
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
                    return _curPHld.AfterAutoSpacing ?? _inheritFromStyleFormat.AfterAutoSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.AfterAutoSpacing;
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
                    return _curPHld.LineSpacing >= 0 ? _curPHld.LineSpacing : _inheritFromStyleFormat.LineSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.LineSpacing;
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
                    return _curPHld.LineSpacingRule != LineSpacingRule.None ? _curPHld.LineSpacingRule : _inheritFromStyleFormat.LineSpacingRule;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.LineSpacingRule;
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
        /// Gets or sets a value indicating whether allow punctuation to overflow boundaries.
        /// </summary>
        public bool OverflowPunctuation
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.OverflowPunctuation ?? _inheritFromStyleFormat.OverflowPunctuation;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.OverflowPunctuation;
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
        /// Gets or sets a value indicating whether allow top line punctuation compression.
        /// </summary>
        public bool TopLinePunctuation
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.TopLinePunctuation ?? _inheritFromStyleFormat.TopLinePunctuation;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.TopLinePunctuation;
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
        /// Gets or sets a value indicating whether the right indentation is automatically adjusted if a document grid is defined.
        /// </summary>
        public bool AdjustRightIndent
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.AdjustRightIndent ?? _inheritFromStyleFormat.AdjustRightIndent;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.AdjustRightIndent;
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
        /// Gets or sets a value indicating whether snap to the grid if a document grid is defined.
        /// </summary>
        public bool SnapToGrid
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return _curPHld.SnapToGrid ?? _inheritFromStyleFormat.SnapToGrid;
                }
                else if (_ownerStyle != null)
                {
                    return _inheritFromBaseStyleFormat.SnapToGrid;
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
        #endregion
    }
}
