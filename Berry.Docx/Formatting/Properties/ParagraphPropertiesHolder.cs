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
        private readonly Document _doc;
        private readonly W.Paragraph _paragraph;
        private readonly W.Style _style;

        // Normal
        private EnumValue<JustificationType> _justificaton;
        private EnumValue<OutlineLevelType> _outlineLevel;
        // Indentation
        private FloatValue _leftInd;
        private FloatValue _leftIndChars;
        private FloatValue _rightInd;
        private FloatValue _rightIndChars;
        private FloatValue _hangingInd;
        private FloatValue _hangingIndChars;
        private FloatValue _firstLineInd;
        private FloatValue _firstLineIndChars;
        private BooleanValue _mirrorIndents;
        private BooleanValue _adjustRightInd;
        // Spacing
        private FloatValue _beforeSpacing;
        private FloatValue _beforeSpacingLines;
        private BooleanValue _beforeAutoSpacing;
        private FloatValue _afterSpacing;
        private FloatValue _afterSpacingLines;
        private BooleanValue _afterAutoSpacing;
        private FloatValue _lineSpacing;
        private EnumValue<LineSpacingRule> _lineSpacingRule;
        private BooleanValue _contextualSpacing;
        private BooleanValue _snapToGrid;
        // Pagination
        private BooleanValue _widowControl;
        private BooleanValue _keepNext;
        private BooleanValue _keepLines;
        private BooleanValue _pageBreakBefore;
        // Formatting Exceptions
        private BooleanValue _suppressLineNumbers;
        private BooleanValue _suppressAutoHyphens;
        // Line Break
        private BooleanValue _kinsoku;
        private BooleanValue _wordWrap;
        private BooleanValue _overflowPunct;
        // Character Spacing
        private BooleanValue _topLinePunct;
        private BooleanValue _autoSpaceDE;
        private BooleanValue _autoSpaceDN;
        private EnumValue<VerticalTextAlignment> _textAlignment;
        #endregion

        #region Constructors
        public ParagraphPropertiesHolder() { }

        public ParagraphPropertiesHolder(Document doc, W.Paragraph paragraph)
        {
            _doc = doc;
            _paragraph = paragraph;
        }

        public ParagraphPropertiesHolder(Document doc, W.Style style)
        {
            _doc = doc;
            _style = style;

            // Numbering
            
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets paragraph numbering format.
        /// </summary>
        public ListFormat ListFormat
        {
            get
            {
                if(_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties?.NumberingProperties?.NumberingId == null) return null;
                    int numId = _paragraph.ParagraphProperties.NumberingProperties.NumberingId.Val;
                    if (_paragraph.ParagraphProperties.NumberingProperties.NumberingLevelReference != null) return null;
                    int ilvl = _paragraph.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val;
                    if (_doc.Package.MainDocumentPart.NumberingDefinitionsPart?.Numbering == null) return null;
                    W.Numbering numbering = _doc.Package.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                    W.NumberingInstance num = numbering.Elements<W.NumberingInstance>().Where(n => n.NumberID == numId).FirstOrDefault();
                    if (num == null) return null;
                    int abstractNumId = num.AbstractNumId.Val;
                    W.AbstractNum abstractNum = numbering.Elements<W.AbstractNum>().Where(a => a.AbstractNumberId == abstractNumId).FirstOrDefault();
                    if (abstractNum == null) return null;
                    return new ListFormat(_doc, abstractNum, ilvl);
                }
                else if(_style != null)
                {
                    if (_style.StyleParagraphProperties?.NumberingProperties?.NumberingId == null) return null;
                    int numId = _style.StyleParagraphProperties.NumberingProperties.NumberingId.Val;
                    string styleId = _style.StyleId;
                    if (_doc.Package.MainDocumentPart.NumberingDefinitionsPart?.Numbering == null) return null;
                    W.Numbering numbering = _doc.Package.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                    W.NumberingInstance num = numbering.Elements<W.NumberingInstance>().Where(n => n.NumberID == numId).FirstOrDefault();
                    if (num == null) return null;
                    int abstractNumId = num.AbstractNumId.Val;
                    W.AbstractNum abstractNum = numbering.Elements<W.AbstractNum>().Where(a => a.AbstractNumberId == abstractNumId).FirstOrDefault();
                    if (abstractNum == null) return null;
                    return new ListFormat(_doc, abstractNum, styleId);
                }
                return null;
            }
        }

        #region Normal
        /// <summary>
        /// Gets or sets the justification.
        /// </summary>
        public EnumValue<JustificationType> Justification
        {
            get
            {
                if(_paragraph != null)
                {
                    W.Justification jc = _paragraph.ParagraphProperties?.Justification;
                    if(jc?.Val == null) return null;
                    return jc.Val.Value.Convert<JustificationType>();
                }
                else if(_style != null)
                {
                    W.Justification jc = _style.StyleParagraphProperties?.Justification;
                    if (jc?.Val == null) return null;
                    return jc.Val.Value.Convert<JustificationType>();
                }
                else
                {
                    return _justificaton;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if(_paragraph.ParagraphProperties.Justification == null)
                    {
                        _paragraph.ParagraphProperties.Justification = new W.Justification();
                    }
                    _paragraph.ParagraphProperties.Justification.Val = value.Val.Convert<W.JustificationValues>();
                }
                else if(_style != null)
                {
                    if(_style.StyleParagraphProperties.Justification == null)
                    {
                        _style.StyleParagraphProperties.Justification= new W.Justification();
                    }
                    _style.StyleParagraphProperties.Justification.Val = value.Val.Convert<W.JustificationValues>();
                }
                else
                {
                    _justificaton = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the outline level.
        /// </summary>
        public EnumValue<OutlineLevelType> OutlineLevel
        {
            get
            {
                if (_paragraph != null)
                {
                    W.OutlineLevel outline = _paragraph.ParagraphProperties?.OutlineLevel;
                    if (outline?.Val == null) return null;
                    return (OutlineLevelType)outline.Val.Value;
                }
                else if (_style != null)
                {
                    W.OutlineLevel outline = _style.StyleParagraphProperties?.OutlineLevel;
                    if (outline?.Val == null) return null;
                    return (OutlineLevelType)outline.Val.Value;
                }
                else
                {
                    return _outlineLevel;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.OutlineLevel == null)
                    {
                        _paragraph.ParagraphProperties.OutlineLevel = new W.OutlineLevel();
                    }
                    _paragraph.ParagraphProperties.OutlineLevel.Val = (int)value.Val;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.OutlineLevel == null)
                    {
                        _style.StyleParagraphProperties.OutlineLevel = new W.OutlineLevel();
                    }
                    _style.StyleParagraphProperties.OutlineLevel.Val = (int)value.Val;
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
        public FloatValue LeftIndent
        {
            get
            {
                W.Indentation ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.Indentation;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.Indentation;
                }
                else
                {
                    return _leftInd;
                }
                if (ele?.Left == null) return null;
                float.TryParse(ele.Left, out float val);
                return val / 20;
            }
            set
            {
                InitIndentation();
                W.Indentation ind;
                if (_paragraph != null)
                {
                    ind = _paragraph.ParagraphProperties.Indentation;
                }
                else if (_style != null)
                {
                    ind = _style.StyleParagraphProperties.Indentation;
                }
                else
                {
                    _leftInd = value;
                    return;
                }
                ind.Left = ((int)(value * 20)).ToString();
            }
        }

        /// <summary>
        /// Gets or sets the right indent (in points) for paragraph.
        /// </summary>

        public FloatValue RightIndent
        {
            get
            {
                W.Indentation ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.Indentation;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.Indentation;
                }
                else
                {
                    return _rightInd;
                }
                if (ele?.Right == null) return null;
                float.TryParse(ele.Right, out float val);
                return val / 20;
            }
            set
            {
                InitIndentation();
                W.Indentation ind;
                if (_paragraph != null)
                {
                    ind = _paragraph.ParagraphProperties.Indentation;
                }
                else if (_style != null)
                {
                    ind = _style.StyleParagraphProperties.Indentation;
                }
                else
                {
                    _rightInd = value;
                    return;
                }
                ind.Right = ((int)(value * 20)).ToString();
            }
        }

        /// <summary>
        /// Gets or sets the left indent (in chars) for paragraph.
        /// </summary>
        public FloatValue LeftCharsIndent
        {
            get
            {
                W.Indentation ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.Indentation;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.Indentation;
                }
                else
                {
                    return _leftIndChars;
                }
                if (ele?.LeftChars == null) return null;
                return ele.LeftChars.Value / 100.0F;
            }
            set
            {
                InitIndentation();
                W.Indentation ind;
                if (_paragraph != null)
                {
                    ind = _paragraph.ParagraphProperties.Indentation;
                }
                else if (_style != null)
                {
                    ind = _style.StyleParagraphProperties.Indentation;
                }
                else
                {
                    _leftIndChars = value;
                    return;
                }
                ind.LeftChars = (int)(value * 100);
            }
        }
        /// <summary>
        /// Gets or sets the right indent (in chars) for paragraph.
        /// </summary>
        public FloatValue RightCharsIndent
        {
            get
            {
                W.Indentation ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.Indentation;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.Indentation;
                }
                else
                {
                    return _rightIndChars;
                }
                if (ele?.RightChars == null) return null;
                return ele.RightChars.Value / 100.0F;
            }
            set
            {
                InitIndentation();
                W.Indentation ind;
                if (_paragraph != null)
                {
                    ind = _paragraph.ParagraphProperties.Indentation;
                }
                else if (_style != null)
                {
                    ind = _style.StyleParagraphProperties.Indentation;
                }
                else
                {
                    _rightIndChars = value;
                    return;
                }
                ind.RightChars = (int)(value * 100);
            }
        }
        /// <summary>
        /// Gets or sets the first line indent (in points) for paragraph.
        /// </summary>
        public FloatValue FirstLineIndent
        {
            get
            {
                W.Indentation ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.Indentation;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.Indentation;
                }
                else
                {
                    return _firstLineInd;
                }
                if (ele?.FirstLine == null) return null;
                float.TryParse(ele.FirstLine, out float val);
                return val / 20;
            }
            set
            {
                InitIndentation();
                W.Indentation ind;
                if (_paragraph != null)
                {
                    ind = _paragraph.ParagraphProperties.Indentation;
                }
                else if (_style != null)
                {
                    ind = _style.StyleParagraphProperties.Indentation;
                }
                else
                {
                    _firstLineInd = value;
                    return;
                }
                if(value != null)
                    ind.FirstLine = ((int)(value * 20)).ToString();
                else
                    ind.FirstLine = null;
            }
        }

        /// <summary>
        /// Gets or sets the first line indent (in chars) for paragraph.
        /// </summary>
        public FloatValue FirstLineCharsIndent
        {
            get
            {
                W.Indentation ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.Indentation;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.Indentation;
                }
                else
                {
                    return _firstLineIndChars;
                }
                if (ele?.FirstLineChars == null) return null;
                return ele.FirstLineChars.Value / 100.0F;
            }
            set
            {
                InitIndentation();
                W.Indentation ind;
                if (_paragraph != null)
                {
                    ind = _paragraph.ParagraphProperties.Indentation;
                }
                else if (_style != null)
                {
                    ind = _style.StyleParagraphProperties.Indentation;
                }
                else
                {
                    _firstLineIndChars = value;
                    return;
                }
                if(value != null)
                    ind.FirstLineChars = (int)(value * 100);
                else 
                    ind.FirstLineChars = null;
            }
        }
        /// <summary>
        /// Gets or sets the hanging indent (in points) for paragraph.
        /// </summary>
        public FloatValue HangingIndent
        {
            get
            {
                W.Indentation ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.Indentation;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.Indentation;
                }
                else
                {
                    return _hangingInd;
                }
                if (ele?.Hanging == null) return null;
                float.TryParse(ele.Hanging, out float val);
                return val / 20;
            }
            set
            {
                InitIndentation();
                W.Indentation ind;
                if (_paragraph != null)
                {
                    ind = _paragraph.ParagraphProperties.Indentation;
                }
                else if (_style != null)
                {
                    ind = _style.StyleParagraphProperties.Indentation;
                }
                else
                {
                    _hangingInd = value;
                    return;
                }
                if (value != null)
                    ind.Hanging = ((int)(value * 20)).ToString();
                else
                    ind.Hanging = null;
            }
        }
        /// <summary>
        /// Gets or sets the hanging indent (in chars) for paragraph.
        /// </summary>
        public FloatValue HangingCharsIndent
        {
            get
            {
                W.Indentation ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.Indentation;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.Indentation;
                }
                else
                {
                    return _hangingIndChars;
                }
                if (ele?.HangingChars == null) return null;
                return ele.HangingChars.Value / 100.0F;
            }
            set
            {
                InitIndentation();
                W.Indentation ind;
                if (_paragraph != null)
                {
                    ind = _paragraph.ParagraphProperties.Indentation;
                }
                else if (_style != null)
                {
                    ind = _style.StyleParagraphProperties.Indentation;
                }
                else
                {
                    _hangingIndChars = value;
                    return;
                }
                if (value != null)
                    ind.HangingChars = (int)(value * 100);
                else
                    ind.HangingChars = null;
            }
        }

        
        /// <summary>
        /// Gets or sets a value indicating whether the paragraph indents should be interpreted as mirrored indents.
        /// </summary>
        public BooleanValue MirrorIndents
        {
            get
            {
                if (_paragraph != null)
                {
                    W.MirrorIndents ele = _paragraph.ParagraphProperties?.MirrorIndents;
                    if (ele == null) return null;
                    if(ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.MirrorIndents ele = _style.StyleParagraphProperties?.MirrorIndents;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _mirrorIndents;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.MirrorIndents == null)
                    {
                        _paragraph.ParagraphProperties.MirrorIndents = new W.MirrorIndents();
                    }
                    if(value)  _paragraph.ParagraphProperties.MirrorIndents.Val = null;
                    else _paragraph.ParagraphProperties.MirrorIndents.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.MirrorIndents == null)
                    {
                        _style.StyleParagraphProperties.MirrorIndents = new W.MirrorIndents();
                    }
                    if (value) _style.StyleParagraphProperties.MirrorIndents.Val = null;
                    else _style.StyleParagraphProperties.MirrorIndents.Val = false;
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
        public BooleanValue AdjustRightIndent
        {
            get
            {
                if (_paragraph != null)
                {
                    W.AdjustRightIndent ele = _paragraph.ParagraphProperties?.AdjustRightIndent;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.AdjustRightIndent ele = _style.StyleParagraphProperties?.AdjustRightIndent;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _adjustRightInd;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.AdjustRightIndent == null)
                    {
                        _paragraph.ParagraphProperties.AdjustRightIndent = new W.AdjustRightIndent();
                    }
                    if (value) _paragraph.ParagraphProperties.AdjustRightIndent.Val = null;
                    else _paragraph.ParagraphProperties.AdjustRightIndent.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.AdjustRightIndent == null)
                    {
                        _style.StyleParagraphProperties.AdjustRightIndent = new W.AdjustRightIndent();
                    }
                    if (value) _style.StyleParagraphProperties.AdjustRightIndent.Val = null;
                    else _style.StyleParagraphProperties.AdjustRightIndent.Val = false; 
                }
                else
                {
                    _adjustRightInd = value;
                }
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
                W.SpacingBetweenLines ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.SpacingBetweenLines;
                }
                else
                {
                    return _beforeSpacing;
                }
                if (ele?.Before == null) return null;
                float.TryParse(ele.Before, out float val);
                return val / 20;
            }
            set
            {
                InitSpacing();
                W.SpacingBetweenLines spacing;
                if (_paragraph != null)
                {
                    spacing = _paragraph.ParagraphProperties.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    spacing = _style.StyleParagraphProperties.SpacingBetweenLines;
                }
                else
                {
                    _beforeSpacing = value;
                    return;
                }
                spacing.Before = ((int)(value * 20)).ToString();
            }
        }

        /// <summary>
        /// Gets or sets the spacing (in lines) before the paragraph.
        /// </summary>
        public FloatValue BeforeLinesSpacing
        {
            get
            {
                W.SpacingBetweenLines ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.SpacingBetweenLines;
                }
                else
                {
                    return _beforeSpacingLines;
                }
                if (ele?.BeforeLines == null) return null;
                return ele.BeforeLines.Value / 100.0F;
            }
            set
            {
                InitSpacing();
                W.SpacingBetweenLines spacing;
                if (_paragraph != null)
                {
                    spacing = _paragraph.ParagraphProperties.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    spacing = _style.StyleParagraphProperties.SpacingBetweenLines;
                }
                else
                {
                    _beforeSpacingLines = value;
                    return;
                }
                spacing.BeforeLines = (int)(value * 100);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether spacing before is automatic.
        /// </summary>
        public BooleanValue BeforeAutoSpacing
        {
            get
            {
                W.SpacingBetweenLines ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.SpacingBetweenLines;
                }
                else
                {
                    return _beforeAutoSpacing;
                }
                if (ele?.BeforeAutoSpacing == null) return null;
                return ele.BeforeAutoSpacing.Value;
            }
            set
            {
                InitSpacing();
                W.SpacingBetweenLines spacing;
                if (_paragraph != null)
                {
                    spacing = _paragraph.ParagraphProperties.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    spacing = _style.StyleParagraphProperties.SpacingBetweenLines;
                }
                else
                {
                    _beforeAutoSpacing = value;
                    return;
                }
                spacing.BeforeAutoSpacing = value.Val;
            }
        }

        /// <summary>
        /// Gets or sets the spacing (in points) after the paragraph.
        /// </summary>
        public FloatValue AfterSpacing
        {
            get
            {
                W.SpacingBetweenLines ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.SpacingBetweenLines;
                }
                else
                {
                    return _afterSpacing;
                }
                if (ele?.After == null) return null;
                float.TryParse(ele.After, out float val);
                return val / 20;
            }
            set
            {
                InitSpacing();
                W.SpacingBetweenLines spacing;
                if (_paragraph != null)
                {
                    spacing = _paragraph.ParagraphProperties.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    spacing = _style.StyleParagraphProperties.SpacingBetweenLines;
                }
                else
                {
                    _afterSpacing = value;
                    return;
                }
                spacing.After = ((int)(value * 20)).ToString();
            }
        }

        /// <summary>
        /// Gets or sets the spacing (in lines) after the paragraph.
        /// </summary>
        public FloatValue AfterLinesSpacing
        {
            get
            {
                W.SpacingBetweenLines ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.SpacingBetweenLines;
                }
                else
                {
                    return _afterSpacingLines;
                }
                if (ele?.AfterLines == null) return null;
                return ele.AfterLines.Value / 100.0F;
            }
            set
            {
                InitSpacing();
                W.SpacingBetweenLines spacing;
                if (_paragraph != null)
                {
                    spacing = _paragraph.ParagraphProperties.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    spacing = _style.StyleParagraphProperties.SpacingBetweenLines;
                }
                else
                {
                    _afterSpacingLines = value;
                    return;
                }
                spacing.AfterLines = (int)(value * 100);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether spacing after is automatic.
        /// </summary>
        public BooleanValue AfterAutoSpacing
        {
            get
            {
                W.SpacingBetweenLines ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.SpacingBetweenLines;
                }
                else
                {
                    return _afterAutoSpacing;
                }
                if (ele?.AfterAutoSpacing == null) return null;
                return ele.AfterAutoSpacing.Value;
            }
            set
            {
                InitSpacing();
                W.SpacingBetweenLines spacing;
                if (_paragraph != null)
                {
                    spacing = _paragraph.ParagraphProperties.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    spacing = _style.StyleParagraphProperties.SpacingBetweenLines;
                }
                else
                {
                    _afterAutoSpacing = value;
                    return;
                }
                spacing.AfterAutoSpacing = value.Val;
            }
        }

        /// <summary>
        /// Gets or sets the line spacing (in points) for paragraph.
        /// </summary>
        public FloatValue LineSpacing
        {
            get
            {
                W.SpacingBetweenLines ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.SpacingBetweenLines;
                }
                else
                {
                    return _lineSpacing;
                }
                if (ele?.Line == null) return null;
                float.TryParse(ele.Line, out float val);
                return val / 20;
            }
            set
            {
                InitSpacing();
                W.SpacingBetweenLines spacing;
                if (_paragraph != null)
                {
                    spacing = _paragraph.ParagraphProperties.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    spacing = _style.StyleParagraphProperties.SpacingBetweenLines;
                }
                else
                {
                    _lineSpacing = value;
                    return;
                }
                spacing.Line = ((int)(value * 20)).ToString();
            }
        }
        /// <summary>
        /// Gets or sets the line spacing rule of paragraph.
        /// </summary>
        public EnumValue<LineSpacingRule> LineSpacingRule
        {
            get
            {
                W.SpacingBetweenLines ele;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    ele = _style.StyleParagraphProperties?.SpacingBetweenLines;
                }
                else
                {
                    return _lineSpacingRule;
                }
                if (ele?.LineRule == null) return null;
                if (ele.LineRule.Value == W.LineSpacingRuleValues.AtLeast)
                    return Docx.LineSpacingRule.AtLeast;
                else if (ele.LineRule.Value == W.LineSpacingRuleValues.Exact)
                    return Docx.LineSpacingRule.Exactly;
                else
                    return Docx.LineSpacingRule.Multiple;
            }
            set
            {
                InitSpacing();
                W.SpacingBetweenLines spacing;
                if (_paragraph != null)
                {
                    spacing = _paragraph.ParagraphProperties.SpacingBetweenLines;
                }
                else if (_style != null)
                {
                    spacing = _style.StyleParagraphProperties.SpacingBetweenLines;
                }
                else
                {
                    _lineSpacingRule = value;
                    return;
                }
                if (value == Docx.LineSpacingRule.AtLeast)
                    spacing.LineRule = W.LineSpacingRuleValues.AtLeast;
                else if (value == Docx.LineSpacingRule.Exactly)
                    spacing.LineRule = W.LineSpacingRuleValues.Exact;
                else if (value == Docx.LineSpacingRule.Multiple)
                    spacing.LineRule = W.LineSpacingRuleValues.Auto;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether don't add space between paragraphs of the same style.
        /// </summary>
        public BooleanValue ContextualSpacing
        {
            get
            {
                if (_paragraph != null)
                {
                    W.ContextualSpacing ele = _paragraph.ParagraphProperties?.ContextualSpacing;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.ContextualSpacing ele = _style.StyleParagraphProperties?.ContextualSpacing;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _contextualSpacing;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.ContextualSpacing == null)
                    {
                        _paragraph.ParagraphProperties.ContextualSpacing = new W.ContextualSpacing();
                    }
                    if (value) _paragraph.ParagraphProperties.ContextualSpacing.Val = null;
                    else _paragraph.ParagraphProperties.ContextualSpacing.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.ContextualSpacing == null)
                    {
                        _style.StyleParagraphProperties.ContextualSpacing = new W.ContextualSpacing();
                    }
                    if (value) _style.StyleParagraphProperties.ContextualSpacing.Val = null;
                    else _style.StyleParagraphProperties.ContextualSpacing.Val = false;
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
        public BooleanValue SnapToGrid
        {
            get
            {
                if (_paragraph != null)
                {
                    W.SnapToGrid ele = _paragraph.ParagraphProperties?.SnapToGrid;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.SnapToGrid ele = _style.StyleParagraphProperties?.SnapToGrid;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _snapToGrid;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.SnapToGrid == null)
                    {
                        _paragraph.ParagraphProperties.SnapToGrid = new W.SnapToGrid();
                    }
                    if (value) _paragraph.ParagraphProperties.SnapToGrid.Val = null;
                    else _paragraph.ParagraphProperties.SnapToGrid.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.SnapToGrid == null)
                    {
                        _style.StyleParagraphProperties.SnapToGrid = new W.SnapToGrid();
                    }
                    if (value) _style.StyleParagraphProperties.SnapToGrid.Val = null;
                    else _style.StyleParagraphProperties.SnapToGrid.Val = false;
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
        public BooleanValue WidowControl
        {
            get
            {
                if (_paragraph != null)
                {
                    W.WidowControl ele = _paragraph.ParagraphProperties?.WidowControl;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.WidowControl ele = _style.StyleParagraphProperties?.WidowControl;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _widowControl;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.WidowControl == null)
                    {
                        _paragraph.ParagraphProperties.WidowControl = new W.WidowControl();
                    }
                    if (value) _paragraph.ParagraphProperties.WidowControl.Val = null;
                    else _paragraph.ParagraphProperties.WidowControl.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.WidowControl == null)
                    {
                        _style.StyleParagraphProperties.WidowControl = new W.WidowControl();
                    }
                    if (value) _style.StyleParagraphProperties.WidowControl.Val = null;
                    else _style.StyleParagraphProperties.WidowControl.Val = false;
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
        public BooleanValue KeepNext
        {
            get
            {
                if (_paragraph != null)
                {
                    W.KeepNext ele = _paragraph.ParagraphProperties?.KeepNext;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.KeepNext ele = _style.StyleParagraphProperties?.KeepNext;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _keepNext;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.KeepNext == null)
                    {
                        _paragraph.ParagraphProperties.KeepNext = new W.KeepNext();
                    }
                    if (value) _paragraph.ParagraphProperties.KeepNext.Val = null;
                    else _paragraph.ParagraphProperties.KeepNext.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.KeepNext == null)
                    {
                        _style.StyleParagraphProperties.KeepNext = new W.KeepNext();
                    }
                    if (value) _style.StyleParagraphProperties.KeepNext.Val = null;
                    else _style.StyleParagraphProperties.KeepNext.Val = false;
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
        public BooleanValue KeepLines
        {
            get
            {
                if (_paragraph != null)
                {
                    W.KeepLines ele = _paragraph.ParagraphProperties?.KeepLines;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.KeepLines ele = _style.StyleParagraphProperties?.KeepLines;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _keepLines;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.KeepLines == null)
                    {
                        _paragraph.ParagraphProperties.KeepLines = new W.KeepLines();
                    }
                    if (value) _paragraph.ParagraphProperties.KeepLines.Val = null;
                    else _paragraph.ParagraphProperties.KeepLines.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.KeepLines == null)
                    {
                        _style.StyleParagraphProperties.KeepLines = new W.KeepLines();
                    }
                    if (value) _style.StyleParagraphProperties.KeepLines.Val = null;
                    else _style.StyleParagraphProperties.KeepLines.Val = false;
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
        public BooleanValue PageBreakBefore
        {
            get
            {
                if (_paragraph != null)
                {
                    W.PageBreakBefore ele = _paragraph.ParagraphProperties?.PageBreakBefore;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.PageBreakBefore ele = _style.StyleParagraphProperties?.PageBreakBefore;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _pageBreakBefore;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.PageBreakBefore == null)
                    {
                        _paragraph.ParagraphProperties.PageBreakBefore = new W.PageBreakBefore();
                    }
                    if (value) _paragraph.ParagraphProperties.PageBreakBefore.Val = null;
                    else _paragraph.ParagraphProperties.PageBreakBefore.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.PageBreakBefore == null)
                    {
                        _style.StyleParagraphProperties.PageBreakBefore = new W.PageBreakBefore();
                    }
                    if (value) _style.StyleParagraphProperties.PageBreakBefore.Val = null;
                    else _style.StyleParagraphProperties.PageBreakBefore.Val = false;
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
        public BooleanValue SuppressLineNumbers
        {
            get
            {
                if (_paragraph != null)
                {
                    W.SuppressLineNumbers ele = _paragraph.ParagraphProperties?.SuppressLineNumbers;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.SuppressLineNumbers ele = _style.StyleParagraphProperties?.SuppressLineNumbers;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _suppressLineNumbers;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.SuppressLineNumbers == null)
                    {
                        _paragraph.ParagraphProperties.SuppressLineNumbers = new W.SuppressLineNumbers();
                    }
                    if (value) _paragraph.ParagraphProperties.SuppressLineNumbers.Val = null;
                    else _paragraph.ParagraphProperties.SuppressLineNumbers.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.SuppressLineNumbers == null)
                    {
                        _style.StyleParagraphProperties.SuppressLineNumbers = new W.SuppressLineNumbers();
                    }
                    if (value) _style.StyleParagraphProperties.SuppressLineNumbers.Val = null;
                    else _style.StyleParagraphProperties.SuppressLineNumbers.Val = false;
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
        public BooleanValue SuppressAutoHyphens
        {
            get
            {
                if (_paragraph != null)
                {
                    W.SuppressAutoHyphens ele = _paragraph.ParagraphProperties?.SuppressAutoHyphens;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.SuppressAutoHyphens ele = _style.StyleParagraphProperties?.SuppressAutoHyphens;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _suppressAutoHyphens;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.SuppressAutoHyphens == null)
                    {
                        _paragraph.ParagraphProperties.SuppressAutoHyphens = new W.SuppressAutoHyphens();
                    }
                    if (value) _paragraph.ParagraphProperties.SuppressAutoHyphens.Val = null;
                    else _paragraph.ParagraphProperties.SuppressAutoHyphens.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.SuppressAutoHyphens == null)
                    {
                        _style.StyleParagraphProperties.SuppressAutoHyphens = new W.SuppressAutoHyphens();
                    }
                    if (value) _style.StyleParagraphProperties.SuppressAutoHyphens.Val = null;
                    else _style.StyleParagraphProperties.SuppressAutoHyphens.Val = false;
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
        public BooleanValue Kinsoku
        {
            get
            {
                if (_paragraph != null)
                {
                    W.Kinsoku ele = _paragraph.ParagraphProperties?.Kinsoku;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.Kinsoku ele = _style.StyleParagraphProperties?.Kinsoku;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _kinsoku;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.Kinsoku == null)
                    {
                        _paragraph.ParagraphProperties.Kinsoku = new W.Kinsoku();
                    }
                    if (value) _paragraph.ParagraphProperties.Kinsoku.Val = null;
                    else _paragraph.ParagraphProperties.Kinsoku.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.Kinsoku == null)
                    {
                        _style.StyleParagraphProperties.Kinsoku = new W.Kinsoku();
                    }
                    if (value) _style.StyleParagraphProperties.Kinsoku.Val = null;
                    else _style.StyleParagraphProperties.Kinsoku.Val = false;
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
        public BooleanValue WordWrap
        {
            get
            {
                if (_paragraph != null)
                {
                    W.WordWrap ele = _paragraph.ParagraphProperties?.WordWrap;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.WordWrap ele = _style.StyleParagraphProperties?.WordWrap;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _wordWrap;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.WordWrap == null)
                    {
                        _paragraph.ParagraphProperties.WordWrap = new W.WordWrap();
                    }
                    if (value) _paragraph.ParagraphProperties.WordWrap.Val = null;
                    else _paragraph.ParagraphProperties.WordWrap.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.WordWrap == null)
                    {
                        _style.StyleParagraphProperties.WordWrap = new W.WordWrap();
                    }
                    if (value) _style.StyleParagraphProperties.WordWrap.Val = null;
                    else _style.StyleParagraphProperties.WordWrap.Val = false;
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
        public BooleanValue OverflowPunctuation
        {
            get
            {
                if (_paragraph != null)
                {
                    W.OverflowPunctuation ele = _paragraph.ParagraphProperties?.OverflowPunctuation;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.OverflowPunctuation ele = _style.StyleParagraphProperties?.OverflowPunctuation;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _overflowPunct;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.OverflowPunctuation == null)
                    {
                        _paragraph.ParagraphProperties.OverflowPunctuation = new W.OverflowPunctuation();
                    }
                    if (value) _paragraph.ParagraphProperties.OverflowPunctuation.Val = null;
                    else _paragraph.ParagraphProperties.OverflowPunctuation.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.OverflowPunctuation == null)
                    {
                        _style.StyleParagraphProperties.OverflowPunctuation = new W.OverflowPunctuation();
                    }
                    if (value) _style.StyleParagraphProperties.OverflowPunctuation.Val = null;
                    else _style.StyleParagraphProperties.OverflowPunctuation.Val = false;
                }
                else
                {
                    _overflowPunct = value;
                }
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
                if (_paragraph != null)
                {
                    W.TopLinePunctuation ele = _paragraph.ParagraphProperties?.TopLinePunctuation;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.TopLinePunctuation ele = _style.StyleParagraphProperties?.TopLinePunctuation;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _topLinePunct;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.TopLinePunctuation == null)
                    {
                        _paragraph.ParagraphProperties.TopLinePunctuation = new W.TopLinePunctuation();
                    }
                    if (value) _paragraph.ParagraphProperties.TopLinePunctuation.Val = null;
                    else _paragraph.ParagraphProperties.TopLinePunctuation.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.TopLinePunctuation == null)
                    {
                        _style.StyleParagraphProperties.TopLinePunctuation = new W.TopLinePunctuation();
                    }
                    if (value) _style.StyleParagraphProperties.TopLinePunctuation.Val = null;
                    else _style.StyleParagraphProperties.TopLinePunctuation.Val = false;
                }
                else
                {
                    _topLinePunct = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether automatically adjust space between Asian and Latin text.
        /// </summary>
        public BooleanValue AutoSpaceDE
        {
            get
            {
                if (_paragraph != null)
                {
                    W.AutoSpaceDE ele = _paragraph.ParagraphProperties?.AutoSpaceDE;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.AutoSpaceDE ele = _style.StyleParagraphProperties?.AutoSpaceDE;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _autoSpaceDE;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.AutoSpaceDE == null)
                    {
                        _paragraph.ParagraphProperties.AutoSpaceDE = new W.AutoSpaceDE();
                    }
                    if (value) _paragraph.ParagraphProperties.AutoSpaceDE.Val = null;
                    else _paragraph.ParagraphProperties.AutoSpaceDE.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.AutoSpaceDE == null)
                    {
                        _style.StyleParagraphProperties.AutoSpaceDE = new W.AutoSpaceDE();
                    }
                    if (value) _style.StyleParagraphProperties.AutoSpaceDE.Val = null;
                    else _style.StyleParagraphProperties.AutoSpaceDE.Val = false;
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
        public BooleanValue AutoSpaceDN
        {
            get
            {
                if (_paragraph != null)
                {
                    W.AutoSpaceDN ele = _paragraph.ParagraphProperties?.AutoSpaceDN;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    W.AutoSpaceDN ele = _style.StyleParagraphProperties?.AutoSpaceDN;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else
                {
                    return _autoSpaceDN;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.AutoSpaceDN == null)
                    {
                        _paragraph.ParagraphProperties.AutoSpaceDN = new W.AutoSpaceDN();
                    }
                    if (value) _paragraph.ParagraphProperties.AutoSpaceDN.Val = null;
                    else _paragraph.ParagraphProperties.AutoSpaceDN.Val = false;
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.AutoSpaceDN == null)
                    {
                        _style.StyleParagraphProperties.AutoSpaceDN = new W.AutoSpaceDN();
                    }
                    if (value) _style.StyleParagraphProperties.AutoSpaceDN.Val = null;
                    else _style.StyleParagraphProperties.AutoSpaceDN.Val = false;
                }
                else
                {
                    _autoSpaceDN = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the vertical alignment of all text on each line displayed within a paragraph.
        /// </summary>
        public EnumValue<VerticalTextAlignment> TextAlignment
        {
            get
            {
                if (_paragraph != null)
                {
                    W.TextAlignment ele = _paragraph.ParagraphProperties?.TextAlignment;
                    if (ele?.Val == null) return null;
                    return ele.Val.Value.Convert<VerticalTextAlignment>();
                }
                else if (_style != null)
                {
                    W.TextAlignment ele = _style.StyleParagraphProperties?.TextAlignment;
                    if (ele?.Val == null) return null;
                    return ele.Val.Value.Convert<VerticalTextAlignment>();
                }
                else
                {
                    return _textAlignment;
                }
            }
            set
            {
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties.TextAlignment == null)
                    {
                        _paragraph.ParagraphProperties.TextAlignment = new W.TextAlignment();
                    }
                    _paragraph.ParagraphProperties.TextAlignment.Val = value.Val.Convert<W.VerticalTextAlignmentValues>();
                }
                else if (_style != null)
                {
                    if (_style.StyleParagraphProperties.TextAlignment == null)
                    {
                        _style.StyleParagraphProperties.TextAlignment = new W.TextAlignment();
                    }
                    _style.StyleParagraphProperties.TextAlignment.Val = value.Val.Convert<W.VerticalTextAlignmentValues>();
                }
                else
                {
                    _textAlignment = value;
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
            if (_paragraph?.ParagraphProperties?.FrameProperties != null)
            {
                _paragraph.ParagraphProperties.FrameProperties = null;
            }
            else if(_style?.StyleParagraphProperties?.FrameProperties != null)
            {
                _style.StyleParagraphProperties.FrameProperties = null;
            }
        }
        #endregion

        #region Private Methods
        private void InitParagraphProperties()
        {
            if(_paragraph != null)
            {
                if (_paragraph.ParagraphProperties == null)
                {
                    _paragraph.ParagraphProperties = new W.ParagraphProperties();
                }
            }
            else if(_style != null)
            {
                if (_style.StyleParagraphProperties == null)
                {
                    _style.StyleParagraphProperties = new W.StyleParagraphProperties();
                }
            }
        }

        private void InitIndentation()
        {
            InitParagraphProperties();
            if(_paragraph != null)
            {
                if (_paragraph.ParagraphProperties.Indentation == null)
                    _paragraph.ParagraphProperties.Indentation = new W.Indentation();
            }
            else if(_style != null)
            {
                if (_style.StyleParagraphProperties.Indentation == null)
                    _style.StyleParagraphProperties.Indentation = new W.Indentation();
            }
            
        }

        private void InitSpacing()
        {
            InitParagraphProperties();
            if (_paragraph != null)
            {
                if (_paragraph.ParagraphProperties.SpacingBetweenLines == null)
                    _paragraph.ParagraphProperties.SpacingBetweenLines = new W.SpacingBetweenLines();
            }
            else if (_style != null)
            {
                if (_style.StyleParagraphProperties.SpacingBetweenLines == null)
                    _style.StyleParagraphProperties.SpacingBetweenLines = new W.SpacingBetweenLines();
            }

        }

        /// <summary>
        /// Returns the paragraph format that specified in the style hierarchy of a style.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="style"> The style</param>
        /// <returns>The paragraph format that specified in the style hierarchy.</returns>
        public static ParagraphPropertiesHolder GetParagraphStyleFormatRecursively(Document doc, W.Style style)
        {
            ParagraphPropertiesHolder format = new ParagraphPropertiesHolder();
            ParagraphPropertiesHolder baseFormat = new ParagraphPropertiesHolder();
            // Gets base style format.
            W.Style baseStyle = style.GetBaseStyle(doc);
            if (baseStyle != null)
                baseFormat = GetParagraphStyleFormatRecursively(doc, baseStyle);

            ParagraphPropertiesHolder curSHld = new ParagraphPropertiesHolder(doc, style);
            // Normal
            format.Justification = curSHld.Justification ?? baseFormat.Justification;
            format.OutlineLevel = curSHld.OutlineLevel ?? baseFormat.OutlineLevel;
            // Indentation
            format.MirrorIndents = curSHld.MirrorIndents ?? baseFormat.MirrorIndents;
            format.AdjustRightIndent = curSHld.AdjustRightIndent ?? baseFormat.AdjustRightIndent;
            // Spacing
            format.BeforeSpacing = curSHld.BeforeSpacing ?? baseFormat.BeforeSpacing;
            format.AfterSpacing = curSHld.AfterSpacing ?? baseFormat.AfterSpacing;
            format.BeforeLinesSpacing = curSHld.BeforeLinesSpacing ?? baseFormat.BeforeLinesSpacing;
            format.AfterLinesSpacing = curSHld.AfterLinesSpacing ?? baseFormat.AfterLinesSpacing;
            format.LineSpacing = curSHld.LineSpacing ?? baseFormat.LineSpacing;
            format.LineSpacingRule = curSHld.LineSpacingRule ?? baseFormat.LineSpacingRule;
            format.BeforeAutoSpacing = curSHld.BeforeAutoSpacing ?? baseFormat.BeforeAutoSpacing;
            format.AfterAutoSpacing = curSHld.AfterAutoSpacing ?? baseFormat.AfterAutoSpacing;
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
            format.TextAlignment = curSHld.TextAlignment ?? baseFormat.TextAlignment;
            // Numbering
            if (curSHld.ListFormat != null)
            {
                format.NumberingFormat = curSHld.ListFormat;
            }
            else if (baseFormat.ListFormat != null)
            {
                format.NumberingFormat = new NumberingFormat(_document, baseFormat.ListFormat, style.StyleId);
            }

            return format;
        }
        #region Indentation
        public static Indentation GetParagraphLeftIndentation(Document doc, W.Paragraph paragraph)
        {
            Indentation charsInd = GetParagraphLeftCharsIndentation(doc, paragraph);
            Indentation pointsInd = GetParagraphLeftPointsIndentation(doc, paragraph);
            SpecialIndentation hangingCharsInd = GetParagraphSpecialCharsIndentation(doc, paragraph);
            SpecialIndentation hangingInd = GetParagraphSpecialPointsIndentation(doc, paragraph);
            if (charsInd != null && charsInd.Val != 0)
            {
                return charsInd;
            }
            else if(hangingCharsInd == null || hangingCharsInd.Val == 0 || hangingCharsInd.Type != SpecialIndentationType.Hanging)
            {
                if (pointsInd != null && pointsInd.Val != 0)
                {
                    if (hangingInd != null && hangingInd.Type == SpecialIndentationType.Hanging)
                        return new Indentation(pointsInd.Val - hangingInd.Val, IndentationUnit.Point);
                    else
                        return pointsInd;
                }
                else if (hangingInd != null && hangingInd.Type == SpecialIndentationType.Hanging)
                {
                    return new Indentation(-hangingInd.Val, IndentationUnit.Point);
                }
            }
            return new Indentation(0, IndentationUnit.Character);
        }

        public static Indentation GetParagraphRightIndentation(Document doc, W.Paragraph paragraph)
        {
            Indentation charsInd = GetParagraphRightCharsIndentation(doc, paragraph);
            Indentation pointsInd = GetParagraphRightPointsIndentation(doc, paragraph);
            if (charsInd != null && charsInd.Val != 0)
                return charsInd;
            else if (pointsInd != null && pointsInd.Val != 0)
                return pointsInd;
            else
                return new Indentation(0, IndentationUnit.Character);
        }

        public static SpecialIndentation GetParagraphSpecialIndentation(Document doc, W.Paragraph paragraph)
        {
            SpecialIndentation charsInd = GetParagraphSpecialCharsIndentation(doc, paragraph);
            SpecialIndentation pointsInd = GetParagraphSpecialPointsIndentation(doc, paragraph);
            if (charsInd != null && charsInd.Val != 0)
                return charsInd;
            else if(pointsInd != null && pointsInd.Val != 0)
                return pointsInd;
            else
                return new SpecialIndentation(SpecialIndentationType.None, 0, IndentationUnit.Character);
        }

        public static Indentation GetStyleLeftIndentation(Document doc, W.Style style)
        {
            Indentation charsInd = GetStyleLeftCharsIndentation(doc, style);
            Indentation pointsInd = GetStyleLeftPointsIndentation(doc, style);
            SpecialIndentation hangingInd = GetStyleSpecialPointsIndentation(doc, style);
            if (charsInd != null && charsInd.Val != 0)
            {
                return charsInd;
            }
            else if (pointsInd != null && pointsInd.Val != 0)
            {
                if (hangingInd != null && hangingInd.Type == SpecialIndentationType.Hanging)
                    return new Indentation(pointsInd.Val - hangingInd.Val, IndentationUnit.Point);
                else
                    return pointsInd;
            }
            else if(hangingInd != null && hangingInd.Type == SpecialIndentationType.Hanging)
            {
                return new Indentation(-hangingInd.Val, IndentationUnit.Point);
            }
            else
            {
                return new Indentation(0, IndentationUnit.Character);
            }
        }

        public static Indentation GetStyleRightIndentation(Document doc, W.Style style)
        {
            Indentation charsInd = GetStyleRightCharsIndentation(doc, style);
            Indentation pointsInd = GetStyleRightPointsIndentation(doc, style);
            if (charsInd != null && charsInd.Val != 0)
                return charsInd;
            else if (pointsInd != null && pointsInd.Val != 0)
                return pointsInd;
            else
                return new Indentation(0, IndentationUnit.Character);
        }

        public static SpecialIndentation GetStyleSpecialIndentation(Document doc, W.Style style)
        {
            SpecialIndentation charsInd = GetStyleSpecialCharsIndentation(doc, style);
            SpecialIndentation pointsInd = GetStyleSpecialPointsIndentation(doc, style);
            if (charsInd != null && charsInd.Val != 0)
                return charsInd;
            else if (pointsInd != null && pointsInd.Val != 0)
                return pointsInd;
            else
                return new SpecialIndentation(SpecialIndentationType.None, 0, IndentationUnit.Character);
        }

        #region Paragraph

        private static Indentation GetParagraphLeftCharsIndentation(Document doc, W.Paragraph paragraph)
        {
            ParagraphPropertiesHolder curPHld = new ParagraphPropertiesHolder(doc, paragraph);
            if (curPHld.LeftCharsIndent != null)
            {
                return new Indentation(curPHld.LeftCharsIndent, IndentationUnit.Character);
            }
            else
            {
                return GetStyleLeftCharsIndentation(doc, paragraph.GetStyle(doc));
            }
        }
        private static Indentation GetParagraphLeftPointsIndentation(Document doc, W.Paragraph paragraph)
        {
            ParagraphPropertiesHolder curPHld = new ParagraphPropertiesHolder(doc, paragraph);
            if (curPHld.LeftIndent != null)
            {
                return new Indentation(curPHld.LeftIndent, IndentationUnit.Point);
            }
            else
            {
                return GetStyleLeftPointsIndentation(doc, paragraph.GetStyle(doc));
            }
        }
        private static Indentation GetParagraphRightCharsIndentation(Document doc, W.Paragraph paragraph)
        {
            ParagraphPropertiesHolder curPHld = new ParagraphPropertiesHolder(doc, paragraph);
            if (curPHld.RightCharsIndent != null)
            {
                return new Indentation(curPHld.RightCharsIndent, IndentationUnit.Character);
            }
            else
            {
                return GetStyleRightCharsIndentation(doc, paragraph.GetStyle(doc));
            }
        }

        private static Indentation GetParagraphRightPointsIndentation(Document doc, W.Paragraph paragraph)
        {
            ParagraphPropertiesHolder curPHld = new ParagraphPropertiesHolder(doc, paragraph);
            if (curPHld.RightIndent != null)
            {
                return new Indentation(curPHld.RightIndent, IndentationUnit.Point);
            }
            else
            {
                return GetStyleRightPointsIndentation(doc, paragraph.GetStyle(doc));
            }
        }
        private static SpecialIndentation GetParagraphSpecialCharsIndentation(Document doc, W.Paragraph paragraph)
        {
            ParagraphPropertiesHolder curPHld = new ParagraphPropertiesHolder(doc, paragraph);
            if (curPHld.HangingCharsIndent != null)
            {
                return new SpecialIndentation(SpecialIndentationType.Hanging, curPHld.HangingCharsIndent, IndentationUnit.Character);
            }
            else if (curPHld.FirstLineCharsIndent != null)
            {
                return new SpecialIndentation(SpecialIndentationType.FirstLine, curPHld.FirstLineCharsIndent, IndentationUnit.Character);
            }
            else
            {
                return GetStyleSpecialCharsIndentation(doc, paragraph.GetStyle(doc));
            }
        }

        public static SpecialIndentation GetParagraphSpecialPointsIndentation(Document doc, W.Paragraph paragraph)
        {
            ParagraphPropertiesHolder curPHld = new ParagraphPropertiesHolder(doc, paragraph);
            if (curPHld.HangingIndent != null)
            {
                return new SpecialIndentation(SpecialIndentationType.Hanging, curPHld.HangingIndent, IndentationUnit.Point);
            }
            else if (curPHld.HangingCharsIndent != null) // HangingIndent = HangingCharsIndent
            {
                return new SpecialIndentation(SpecialIndentationType.Hanging, curPHld.HangingCharsIndent * 5, IndentationUnit.Point);
            }
            else if (curPHld.FirstLineIndent != null)
            {
                return new SpecialIndentation(SpecialIndentationType.FirstLine, curPHld.FirstLineIndent, IndentationUnit.Point);
            }
            else if (curPHld.FirstLineCharsIndent != null) // FirstLine = FirstLineCharsIndent
            {
                return new SpecialIndentation(SpecialIndentationType.FirstLine, curPHld.FirstLineCharsIndent * 5, IndentationUnit.Point);
            }
            else
            {
                return GetStyleSpecialPointsIndentation(doc, paragraph.GetStyle(doc));
            }
        }
        #endregion

        #region Style
        private static Indentation GetStyleLeftCharsIndentation(Document doc, W.Style style)
        {
            ParagraphPropertiesHolder curSHld = new ParagraphPropertiesHolder(doc, style);
            if (curSHld.LeftCharsIndent != null)
            {
                return new Indentation(curSHld.LeftCharsIndent, IndentationUnit.Character);
            }
            else
            {
                W.Style baseStyle = style.GetBaseStyle(doc);
                if (baseStyle != null)
                {
                    return GetStyleLeftCharsIndentation(doc, baseStyle);
                }
            }
            return null;
        }

        private static Indentation GetStyleLeftPointsIndentation(Document doc, W.Style style)
        {
            ParagraphPropertiesHolder curSHld = new ParagraphPropertiesHolder(doc, style);
            if (curSHld.LeftIndent != null)
            {
                return new Indentation(curSHld.LeftIndent, IndentationUnit.Point);
            }
            else
            {
                W.Style baseStyle = style.GetBaseStyle(doc);
                if (baseStyle != null)
                {
                    return GetStyleLeftPointsIndentation(doc, baseStyle);
                }
            }
            return null;
        }

        private static Indentation GetStyleRightCharsIndentation(Document doc, W.Style style)
        {
            ParagraphPropertiesHolder curSHld = new ParagraphPropertiesHolder(doc, style);
            if (curSHld.RightCharsIndent != null)
            {
                return new Indentation(curSHld.RightCharsIndent, IndentationUnit.Character);
            }
            else
            {
                W.Style baseStyle = style.GetBaseStyle(doc);
                if (baseStyle != null)
                {
                    return GetStyleRightCharsIndentation(doc, baseStyle);
                }
            }
            return null;
        }

        private static Indentation GetStyleRightPointsIndentation(Document doc, W.Style style)
        {
            ParagraphPropertiesHolder curSHld = new ParagraphPropertiesHolder(doc, style);
            if (curSHld.RightIndent != null)
            {
                return new Indentation(curSHld.RightIndent, IndentationUnit.Point);
            }
            else
            {
                W.Style baseStyle = style.GetBaseStyle(doc);
                if (baseStyle != null)
                {
                    return GetStyleRightPointsIndentation(doc, baseStyle);
                }
            }
            return null;
        }

        private static SpecialIndentation GetStyleSpecialCharsIndentation(Document doc, W.Style style)
        {
            ParagraphPropertiesHolder curSHld = new ParagraphPropertiesHolder(doc, style);
            if (curSHld.HangingCharsIndent != null)
            {
                return new SpecialIndentation(SpecialIndentationType.Hanging, curSHld.HangingCharsIndent, IndentationUnit.Character);
            }
            else if(curSHld.FirstLineCharsIndent != null)
            {
                return new SpecialIndentation(SpecialIndentationType.FirstLine, curSHld.FirstLineCharsIndent, IndentationUnit.Character);
            }
            else
            {
                W.Style baseStyle = style.GetBaseStyle(doc);
                if (baseStyle != null)
                {
                    return GetStyleSpecialCharsIndentation(doc, baseStyle);
                }
            }
            return null;
        }

        public static SpecialIndentation GetStyleSpecialPointsIndentation(Document doc, W.Style style)
        {
            ParagraphPropertiesHolder curSHld = new ParagraphPropertiesHolder(doc, style);
            if (curSHld.HangingIndent != null)
            {
                return new SpecialIndentation(SpecialIndentationType.Hanging, curSHld.HangingIndent, IndentationUnit.Point);
            }
            else if(curSHld.HangingCharsIndent != null) //HangingIndent 的值与 HangingCharsIndent 保持一致
            {
                return new SpecialIndentation(SpecialIndentationType.Hanging, curSHld.HangingCharsIndent * 5, IndentationUnit.Point);
            }
            else if (curSHld.FirstLineIndent != null)
            {
                return new SpecialIndentation(SpecialIndentationType.FirstLine, curSHld.FirstLineIndent, IndentationUnit.Point);
            }
            else if (curSHld.FirstLineCharsIndent != null) //FirstLine 的值与 FirstLineCharsIndent 保持一致
            {
                return new SpecialIndentation(SpecialIndentationType.FirstLine, curSHld.FirstLineCharsIndent * 5, IndentationUnit.Point);
            }
            else
            {
                W.Style baseStyle = style.GetBaseStyle(doc);
                if (baseStyle != null)
                {
                    return GetStyleSpecialPointsIndentation(doc, baseStyle);
                }
            }
            return null;
        }
        #endregion

        #endregion

        #endregion
    }
}
