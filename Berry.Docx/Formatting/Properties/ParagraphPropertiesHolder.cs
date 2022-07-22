using System;
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
        private readonly EnumValue<TableRegionType> _tableStyleRegion;

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
        }

        public ParagraphPropertiesHolder(Document doc, W.Style style, TableRegionType type)
        {
            _doc = doc;
            _style = style;
            _tableStyleRegion = type;
        }
        #endregion

        #region Public Properties

        #region Normal
        /// <summary>
        /// Gets or sets the justification.
        /// </summary>
        public EnumValue<JustificationType> Justification
        {
            get
            {
                if (NoInstance()) return _justificaton;
                W.Justification jc = null;
                if (_paragraph != null)
                {
                    jc = _paragraph.ParagraphProperties?.Justification;
                }
                else if(_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        jc = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.StyleParagraphProperties?.Justification;
                    }
                    else
                    {
                        jc = _style.StyleParagraphProperties?.Justification;
                    }
                }
                if (jc?.Val == null) return null;
                return jc.Val.Value.Convert<JustificationType>();
            }
            set
            {
                _justificaton = value;
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    _paragraph.ParagraphProperties.Justification = new W.Justification() { Val = value.Val.Convert<W.JustificationValues>() };
                }
                else if(_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        tblStylePr.StyleParagraphProperties.Justification = new W.Justification() { Val = value.Val.Convert<W.JustificationValues>() };
                    }
                    else
                    {
                        _style.StyleParagraphProperties.Justification = new W.Justification() { Val = value.Val.Convert<W.JustificationValues>() };
                    }
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
                if (NoInstance()) return _outlineLevel;
                W.OutlineLevel outline = null;
                if (_paragraph != null)
                {
                    outline = _paragraph.ParagraphProperties?.OutlineLevel;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        outline = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.StyleParagraphProperties?.OutlineLevel;
                    }
                    else
                    {
                        outline = _style.StyleParagraphProperties?.OutlineLevel;
                    }
                }
                if (outline?.Val == null) return null;
                return (OutlineLevelType)outline.Val.Value;
            }
            set
            {
                _outlineLevel = value;
                InitParagraphProperties();
                if (_paragraph != null)
                {
                    _paragraph.ParagraphProperties.OutlineLevel = new W.OutlineLevel() { Val = (int)value.Val };
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        tblStylePr.StyleParagraphProperties.OutlineLevel = new W.OutlineLevel() { Val = (int)value.Val };
                    }
                    else
                    {
                        _style.StyleParagraphProperties.OutlineLevel = new W.OutlineLevel() { Val = (int)value.Val };
                    }
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
                if (NoInstance()) return _leftInd;
                TryGetIndentation(out W.Indentation ind);
                if (ind?.Left == null) return null;
                float.TryParse(ind.Left, out float val);
                return val / 20;
            }
            set
            {
                _leftInd = value;
                CreateIndentation();
                TryGetIndentation(out W.Indentation ind);
                if (ind == null) return;
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
                if (NoInstance()) return _rightInd;
                TryGetIndentation(out W.Indentation ind);
                if (ind?.Right == null) return null;
                float.TryParse(ind.Right, out float val);
                return val / 20;
            }
            set
            {
                _rightInd = value;
                CreateIndentation();
                TryGetIndentation(out W.Indentation ind);
                if (ind == null) return;
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
                if (NoInstance()) return _leftIndChars;
                TryGetIndentation(out W.Indentation ind);
                if (ind?.LeftChars == null) return null;
                return ind.LeftChars.Value / 100.0F;
            }
            set
            {
                _leftIndChars = value;
                CreateIndentation();
                TryGetIndentation(out W.Indentation ind);
                if (ind == null) return;
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
                if (NoInstance()) return _rightIndChars;
                TryGetIndentation(out W.Indentation ind);
                if (ind?.RightChars == null) return null;
                return ind.RightChars.Value / 100.0F;
            }
            set
            {
                _rightIndChars = value;
                CreateIndentation();
                TryGetIndentation(out W.Indentation ind);
                if (ind == null) return;
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
                if (NoInstance()) return _firstLineInd;
                TryGetIndentation(out W.Indentation ind);
                if (ind?.FirstLine == null) return null;
                float.TryParse(ind.FirstLine, out float val);
                return val / 20;
            }
            set
            {
                _firstLineInd = value;
                CreateIndentation();
                TryGetIndentation(out W.Indentation ind);
                if (ind == null) return;
                if (value != null)
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
                if (NoInstance()) return _firstLineIndChars;
                TryGetIndentation(out W.Indentation ind);
                if (ind?.FirstLineChars == null) return null;
                return ind.FirstLineChars.Value / 100.0F;
            }
            set
            {
                _firstLineIndChars = value;
                CreateIndentation();
                TryGetIndentation(out W.Indentation ind);
                if (ind == null) return;
                if (value != null)
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
                if (NoInstance()) return _hangingInd;
                TryGetIndentation(out W.Indentation ind);
                if (ind?.Hanging == null) return null;
                float.TryParse(ind.Hanging, out float val);
                return val / 20;
            }
            set
            {
                _hangingInd = value;
                CreateIndentation();
                TryGetIndentation(out W.Indentation ind);
                if (ind == null) return;
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
                if (NoInstance()) return _hangingIndChars;
                TryGetIndentation(out W.Indentation ind);
                if (ind?.HangingChars == null) return null;
                return ind.HangingChars.Value / 100.0F;
            }
            set
            {
                _hangingIndChars = value;
                CreateIndentation();
                TryGetIndentation(out W.Indentation ind);
                if (ind == null) return;
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
                if (NoInstance()) return _mirrorIndents;
                W.MirrorIndents ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.MirrorIndents;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.StyleParagraphProperties?.MirrorIndents;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.MirrorIndents;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _mirrorIndents = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if(value) tblStylePr.StyleParagraphProperties.MirrorIndents = new W.MirrorIndents() { Val = null };
                        else tblStylePr.StyleParagraphProperties.MirrorIndents = new W.MirrorIndents() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.MirrorIndents = new W.MirrorIndents() { Val = null };
                        else _style.StyleParagraphProperties.MirrorIndents = new W.MirrorIndents() { Val = false };
                    }
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
                if (NoInstance()) return _adjustRightInd;
                W.AdjustRightIndent ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.AdjustRightIndent;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.StyleParagraphProperties?.AdjustRightIndent;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.AdjustRightIndent;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _adjustRightInd = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.AdjustRightIndent = new W.AdjustRightIndent() { Val = null };
                        else tblStylePr.StyleParagraphProperties.AdjustRightIndent = new W.AdjustRightIndent() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.AdjustRightIndent = new W.AdjustRightIndent() { Val = null };
                        else _style.StyleParagraphProperties.AdjustRightIndent = new W.AdjustRightIndent() { Val = false };
                    }
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
                if(NoInstance()) return _beforeSpacing;
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing?.Before == null) return null;
                float.TryParse(spacing.Before, out float val);
                return val / 20;
            }
            set
            {
                _beforeSpacing = value;
                CreateSpacing();
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing == null) return;
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
                if (NoInstance()) return _beforeSpacingLines;
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing?.BeforeLines == null) return null;
                return spacing.BeforeLines.Value / 100.0F;
            }
            set
            {
                _beforeSpacingLines = value;
                CreateSpacing();
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing == null) return;
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
                if (NoInstance()) return _beforeAutoSpacing;
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing?.BeforeAutoSpacing == null) return null;
                return spacing.BeforeAutoSpacing.Value;
            }
            set
            {
                _beforeAutoSpacing = value;
                CreateSpacing();
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing == null) return;
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
                if (NoInstance()) return _afterSpacing;
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing?.After == null) return null;
                float.TryParse(spacing.After, out float val);
                return val / 20;
            }
            set
            {
                _afterSpacing = value;
                CreateSpacing();
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing == null) return;
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
                if (NoInstance()) return _afterSpacingLines;
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing?.AfterLines == null) return null;
                return spacing.AfterLines.Value / 100.0F;
            }
            set
            {
                _afterSpacingLines = value;
                CreateSpacing();
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing == null) return;
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
                if (NoInstance()) return _afterAutoSpacing;
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing?.AfterAutoSpacing == null) return null;
                return spacing.AfterAutoSpacing.Value;
            }
            set
            {
                _afterAutoSpacing = value;
                CreateSpacing();
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing == null) return;
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
                if (NoInstance()) return _lineSpacing;
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing?.Line == null) return null;
                float.TryParse(spacing.Line, out float val);
                return val / 20;
            }
            set
            {
                _lineSpacing = value;
                CreateSpacing();
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing == null) return;
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
                if (NoInstance()) return _lineSpacingRule;
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing?.LineRule == null) return null;
                if (spacing.LineRule.Value == W.LineSpacingRuleValues.AtLeast)
                    return Docx.LineSpacingRule.AtLeast;
                else if (spacing.LineRule.Value == W.LineSpacingRuleValues.Exact)
                    return Docx.LineSpacingRule.Exactly;
                else
                    return Docx.LineSpacingRule.Multiple;
            }
            set
            {
                _lineSpacingRule = value;
                CreateSpacing();
                TryGetSpacing(out W.SpacingBetweenLines spacing);
                if (spacing == null) return;
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
                if(NoInstance()) return _contextualSpacing;
                W.ContextualSpacing ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.ContextualSpacing;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.StyleParagraphProperties?.ContextualSpacing;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.ContextualSpacing;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _contextualSpacing = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.ContextualSpacing = new W.ContextualSpacing() { Val = null };
                        else tblStylePr.StyleParagraphProperties.ContextualSpacing = new W.ContextualSpacing() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.ContextualSpacing = new W.ContextualSpacing() { Val = null };
                        else _style.StyleParagraphProperties.ContextualSpacing = new W.ContextualSpacing() { Val = false };
                    }
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
                if(NoInstance()) return _snapToGrid;
                W.SnapToGrid ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.SnapToGrid;
                    if (ele == null) return null;
                    if (ele.Val == null) return true;
                    return ele.Val.Value;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault()
                                ?.StyleParagraphProperties?.SnapToGrid;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.SnapToGrid;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _snapToGrid = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.SnapToGrid = new W.SnapToGrid() { Val = null };
                        else tblStylePr.StyleParagraphProperties.SnapToGrid = new W.SnapToGrid() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.SnapToGrid = new W.SnapToGrid() { Val = null };
                        else _style.StyleParagraphProperties.SnapToGrid = new W.SnapToGrid() { Val = false };
                    }
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
                if (NoInstance()) return _widowControl;
                W.WidowControl ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.WidowControl;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault()
                                ?.StyleParagraphProperties?.WidowControl;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.WidowControl;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _widowControl = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.WidowControl = new W.WidowControl() { Val = null };
                        else tblStylePr.StyleParagraphProperties.WidowControl = new W.WidowControl() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.WidowControl = new W.WidowControl() { Val = null };
                        else _style.StyleParagraphProperties.WidowControl = new W.WidowControl() { Val = false };
                    }
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
                if (NoInstance()) return _keepNext;
                W.KeepNext ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.KeepNext;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault()
                                ?.StyleParagraphProperties?.KeepNext;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.KeepNext;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _keepNext = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.KeepNext = new W.KeepNext() { Val = null };
                        else tblStylePr.StyleParagraphProperties.KeepNext = new W.KeepNext() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.KeepNext = new W.KeepNext() { Val = null };
                        else _style.StyleParagraphProperties.KeepNext = new W.KeepNext() { Val = false };
                    }
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
                if(NoInstance()) return _keepLines;
                W.KeepLines ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.KeepLines;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault()
                                ?.StyleParagraphProperties?.KeepLines;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.KeepLines;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _keepLines = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.KeepLines = new W.KeepLines() { Val = null };
                        else tblStylePr.StyleParagraphProperties.KeepLines = new W.KeepLines() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.KeepLines = new W.KeepLines() { Val = null };
                        else _style.StyleParagraphProperties.KeepLines = new W.KeepLines() { Val = false };
                    }
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
                if(NoInstance()) return _pageBreakBefore;
                W.PageBreakBefore ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.PageBreakBefore;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault()
                                ?.StyleParagraphProperties?.PageBreakBefore;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.PageBreakBefore;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _pageBreakBefore = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.PageBreakBefore = new W.PageBreakBefore() { Val = null };
                        else tblStylePr.StyleParagraphProperties.PageBreakBefore = new W.PageBreakBefore() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.PageBreakBefore = new W.PageBreakBefore() { Val = null };
                        else _style.StyleParagraphProperties.PageBreakBefore = new W.PageBreakBefore() { Val = false };
                    }
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
                if(NoInstance()) return _suppressLineNumbers;
                W.SuppressLineNumbers ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.SuppressLineNumbers;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault()
                                ?.StyleParagraphProperties?.SuppressLineNumbers;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.SuppressLineNumbers;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _suppressLineNumbers = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.SuppressLineNumbers = new W.SuppressLineNumbers() { Val = null };
                        else tblStylePr.StyleParagraphProperties.SuppressLineNumbers = new W.SuppressLineNumbers() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.SuppressLineNumbers = new W.SuppressLineNumbers() { Val = null };
                        else _style.StyleParagraphProperties.SuppressLineNumbers = new W.SuppressLineNumbers() { Val = false };
                    }
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
                if(NoInstance()) return _suppressAutoHyphens;
                W.SuppressAutoHyphens ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.SuppressAutoHyphens;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault()
                                ?.StyleParagraphProperties?.SuppressAutoHyphens;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.SuppressAutoHyphens;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _suppressAutoHyphens = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.SuppressAutoHyphens = new W.SuppressAutoHyphens() { Val = null };
                        else tblStylePr.StyleParagraphProperties.SuppressAutoHyphens = new W.SuppressAutoHyphens() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.SuppressAutoHyphens = new W.SuppressAutoHyphens() { Val = null };
                        else _style.StyleParagraphProperties.SuppressAutoHyphens = new W.SuppressAutoHyphens() { Val = false };
                    }
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
                if(NoInstance()) return _kinsoku;
                W.Kinsoku ele = null;
                if (_paragraph != null)
                {
                   ele = _paragraph.ParagraphProperties?.Kinsoku;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault()
                                ?.StyleParagraphProperties?.Kinsoku;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.Kinsoku;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _kinsoku = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.Kinsoku = new W.Kinsoku() { Val = null };
                        else tblStylePr.StyleParagraphProperties.Kinsoku = new W.Kinsoku() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.Kinsoku = new W.Kinsoku() { Val = null };
                        else _style.StyleParagraphProperties.Kinsoku = new W.Kinsoku() { Val = false };
                    }
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
                if(NoInstance()) return _wordWrap;
                W.WordWrap ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.WordWrap;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault()
                                ?.StyleParagraphProperties?.WordWrap;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.WordWrap;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _wordWrap = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.WordWrap = new W.WordWrap() { Val = null };
                        else tblStylePr.StyleParagraphProperties.WordWrap = new W.WordWrap() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.WordWrap = new W.WordWrap() { Val = null };
                        else _style.StyleParagraphProperties.WordWrap = new W.WordWrap() { Val = false };
                    }
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
                if(NoInstance()) return _overflowPunct;
                W.OverflowPunctuation ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.OverflowPunctuation;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault()
                                ?.StyleParagraphProperties?.OverflowPunctuation;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.OverflowPunctuation;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _overflowPunct = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.OverflowPunctuation = new W.OverflowPunctuation() { Val = null };
                        else tblStylePr.StyleParagraphProperties.OverflowPunctuation = new W.OverflowPunctuation() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.OverflowPunctuation = new W.OverflowPunctuation() { Val = null };
                        else _style.StyleParagraphProperties.OverflowPunctuation = new W.OverflowPunctuation() { Val = false };
                    }
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
                if(NoInstance()) return _topLinePunct;
                W.TopLinePunctuation ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.TopLinePunctuation;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault()
                                ?.StyleParagraphProperties?.TopLinePunctuation;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.TopLinePunctuation;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _topLinePunct = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.TopLinePunctuation = new W.TopLinePunctuation() { Val = null };
                        else tblStylePr.StyleParagraphProperties.TopLinePunctuation = new W.TopLinePunctuation() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.TopLinePunctuation = new W.TopLinePunctuation() { Val = null };
                        else _style.StyleParagraphProperties.TopLinePunctuation = new W.TopLinePunctuation() { Val = false };
                    }
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
                if(NoInstance()) return _autoSpaceDE;
                W.AutoSpaceDE ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.AutoSpaceDE;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault()
                                ?.StyleParagraphProperties?.AutoSpaceDE;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.AutoSpaceDE;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _autoSpaceDE = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.AutoSpaceDE = new W.AutoSpaceDE() { Val = null };
                        else tblStylePr.StyleParagraphProperties.AutoSpaceDE = new W.AutoSpaceDE() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.AutoSpaceDE = new W.AutoSpaceDE() { Val = null };
                        else _style.StyleParagraphProperties.AutoSpaceDE = new W.AutoSpaceDE() { Val = false };
                    }
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
                if(NoInstance()) return _autoSpaceDN;
                W.AutoSpaceDN ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.AutoSpaceDN;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault()
                                ?.StyleParagraphProperties?.AutoSpaceDN;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.AutoSpaceDN;
                    }
                }
                if (ele == null) return null;
                if (ele.Val == null) return true;
                return ele.Val.Value;
            }
            set
            {
                _autoSpaceDN = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (value) tblStylePr.StyleParagraphProperties.AutoSpaceDN = new W.AutoSpaceDN() { Val = null };
                        else tblStylePr.StyleParagraphProperties.AutoSpaceDN = new W.AutoSpaceDN() { Val = false };
                    }
                    else
                    {
                        if (value) _style.StyleParagraphProperties.AutoSpaceDN = new W.AutoSpaceDN() { Val = null };
                        else _style.StyleParagraphProperties.AutoSpaceDN = new W.AutoSpaceDN() { Val = false };
                    }
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
                if(NoInstance()) return _textAlignment;
                W.TextAlignment ele = null;
                if (_paragraph != null)
                {
                    ele = _paragraph.ParagraphProperties?.TextAlignment;
                    if (ele?.Val == null) return null;
                    return ele.Val.Value.Convert<VerticalTextAlignment>();
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        ele = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault()
                                ?.StyleParagraphProperties?.TextAlignment;
                    }
                    else
                    {
                        ele = _style.StyleParagraphProperties?.TextAlignment;
                    }
                }
                if (ele?.Val == null) return null;
                return ele.Val.Value.Convert<VerticalTextAlignment>();
            }
            set
            {
                _textAlignment = value;
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
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        tblStylePr.StyleParagraphProperties.TextAlignment = new W.TextAlignment()
                        {
                            Val = value.Val.Convert<W.VerticalTextAlignmentValues>()
                        };
                    }
                    else
                    {
                        _style.StyleParagraphProperties.TextAlignment = new W.TextAlignment()
                        {
                            Val = value.Val.Convert<W.VerticalTextAlignmentValues>()
                        };
                    }
                }
            }
        }
        #endregion

        #endregion

        #region Public Methods
        /// <summary>
        /// Clears all paragraph formats.
        /// </summary>
        public void ClearFormatting()
        {
            if (_paragraph?.ParagraphProperties != null)
            {
                _paragraph.ParagraphProperties = null;
            }
            else if (_style != null)
            {
                if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                {
                    W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                    W.TableStyleProperties tblPr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                    if(tblPr?.StyleParagraphProperties != null) tblPr.StyleParagraphProperties = null;
                }
                else
                {
                    _style.StyleRunProperties.RemoveAllChildren();
                }
            }
        }

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

        private bool NoInstance()
        {
            return _paragraph == null && _style == null;
        }
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
                if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                {
                    W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                    if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                    {
                        _style.Append(new W.TableStyleProperties() { Type = type });
                    }
                    W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                    if (tblStylePr.StyleParagraphProperties == null)
                    {
                        tblStylePr.StyleParagraphProperties = new W.StyleParagraphProperties();
                    }
                }
                else if (_style.StyleParagraphProperties == null)
                {
                    _style.StyleParagraphProperties = new W.StyleParagraphProperties();
                }
            }
        }

        private void TryGetIndentation(out W.Indentation ind)
        {
            ind = null;
            if (_paragraph != null)
            {
                ind = _paragraph.ParagraphProperties?.Indentation;
            }
            else if (_style != null)
            {
                if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                {
                    ind = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                            ?.StyleParagraphProperties?.Indentation;
                }
                else
                {
                    ind = _style.StyleParagraphProperties?.Indentation;
                }
            }
        }
        private void CreateIndentation()
        {
            InitParagraphProperties();
            if(_paragraph != null)
            {
                if (_paragraph.ParagraphProperties.Indentation == null)
                    _paragraph.ParagraphProperties.Indentation = new W.Indentation();
            }
            else if(_style != null)
            {
                if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                {
                    W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                    W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                    if (tblStylePr.StyleParagraphProperties.Indentation == null)
                    {
                        tblStylePr.StyleParagraphProperties.Indentation = new W.Indentation();
                    }
                }
                else if (_style.StyleParagraphProperties.Indentation == null)
                {
                    _style.StyleParagraphProperties.Indentation = new W.Indentation();
                }
            }
        }

        private void TryGetSpacing(out W.SpacingBetweenLines spacing)
        {
            spacing = null;
            if (_paragraph != null)
            {
                spacing = _paragraph.ParagraphProperties?.SpacingBetweenLines;
            }
            else if (_style != null)
            {
                if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                {
                    spacing = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                            ?.StyleParagraphProperties?.SpacingBetweenLines;
                }
                else
                {
                    spacing = _style.StyleParagraphProperties?.SpacingBetweenLines;
                }
            }
        }

        private void CreateSpacing()
        {
            InitParagraphProperties();
            if (_paragraph != null)
            {
                if (_paragraph.ParagraphProperties.SpacingBetweenLines == null)
                    _paragraph.ParagraphProperties.SpacingBetweenLines = new W.SpacingBetweenLines();
            }
            else if (_style != null)
            {
                if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                {
                    W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                    W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                    if (tblStylePr.StyleParagraphProperties.SpacingBetweenLines == null)
                    {
                        tblStylePr.StyleParagraphProperties.SpacingBetweenLines = new W.SpacingBetweenLines();
                    }
                }
                else if (_style.StyleParagraphProperties.SpacingBetweenLines == null)
                {
                    _style.StyleParagraphProperties.SpacingBetweenLines = new W.SpacingBetweenLines();
                }
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
