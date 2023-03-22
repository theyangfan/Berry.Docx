using System;
using System.Linq;
using System.Drawing;
using P = DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent an OpenXML RunProperties holder.
    /// </summary>
    internal class RunPropertiesHolder
    {
        #region Private Members
        private readonly P.WordprocessingDocument _document;
        private readonly W.Run _run;
        private readonly W.Style _style;
        private readonly EnumValue<TableRegionType> _tableStyleRegion;
        private readonly W.RunPropertiesDefault _defaultRPr;
        private readonly W.Paragraph _paragraph;
        private readonly W.Level _numberingLevel;

        private string _fontNameEastAsia;
        private string _fontNameAscii;
        private string _fontNameHAnsi;
        private string _fontNameCs;
        private EnumValue<FontContentType> _fontTypeHint;
        private FloatValue _fontSize;
        private FloatValue _fontSizeCs;
        private BooleanValue _bold;
        private BooleanValue _italic;
        private EnumValue<SubSuperScript> _subSuperScript;
        private EnumValue<UnderlineStyle> _underlineStyle;
        private ColorValue _color = ColorValue.Auto;
        private IntegerValue _characterScale;
        private FloatValue _characterSpacing;
        private FloatValue _position;
        private BooleanValue _isHidden;
        #endregion

        #region Constructors
        public RunPropertiesHolder() { }
        /// <summary>
        /// Run 元素下的字符属性。
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="run"></param>
        public RunPropertiesHolder(P.WordprocessingDocument doc, W.Run run)
        {
            _document = doc;
            _run = run;
        }

        /// <summary>
        /// 样式的字符属性。
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="style"></param>
        public RunPropertiesHolder(P.WordprocessingDocument doc, W.Style style)
        {
            _document = doc;
            _style = style;
        }

        public RunPropertiesHolder(P.WordprocessingDocument doc, W.Style style, TableRegionType type)
        {
            _document = doc;
            _style = style;
            _tableStyleRegion = type;
        }

        /// <summary>
        /// 文档默认字符属性。
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="rPrDefault"></param>
        public RunPropertiesHolder(P.WordprocessingDocument doc, W.RunPropertiesDefault rPrDefault)
        {
            _document = doc;
            _defaultRPr = rPrDefault;
        }

        /// <summary>
        /// 段落标记的字符属性。
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="paragraph"></param>
        public RunPropertiesHolder(P.WordprocessingDocument doc, W.Paragraph paragraph)
        {
            _document = doc;
            _paragraph = paragraph;
        }

        public RunPropertiesHolder(P.WordprocessingDocument doc, W.Level numberingLevel)
        {
            _document = doc;
            _numberingLevel = numberingLevel;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets or sets East Asian font name.
        /// </summary>
        public string FontNameEastAsia
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _fontNameEastAsia;
                }
                W.RunFonts rFonts = null;
                if(_run?.RunProperties?.RunFonts != null)
                {
                    rFonts = _run.RunProperties.RunFonts;
                }
                else if(_style != null)
                {
                    if(_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        rFonts = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.RunFonts;
                    }
                    else
                    {
                        rFonts = _style.StyleRunProperties?.RunFonts;
                    }
                }
                else if(_defaultRPr?.RunPropertiesBaseStyle?.RunFonts != null)
                {
                    rFonts = _defaultRPr.RunPropertiesBaseStyle.RunFonts;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.RunFonts>() != null)
                {
                    rFonts = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>();
                }
                else if(_numberingLevel?.NumberingSymbolRunProperties?.RunFonts != null)
                {
                    rFonts = _numberingLevel.NumberingSymbolRunProperties.RunFonts;
                }
                if(rFonts?.EastAsiaTheme != null)
                {
                    return _document.GetThemeFont(rFonts.EastAsiaTheme);
                }
                return rFonts?.EastAsia;
            }
            set
            {
                if(_run != null)
                {
                    if(_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if(_run.RunProperties.RunFonts == null)
                    {
                        _run.RunProperties.RunFonts = new W.RunFonts();
                    }
                    _run.RunProperties.RunFonts.EastAsia = value;
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
                        if(tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if(tblStylePr.RunPropertiesBaseStyle.RunFonts == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.RunFonts = new W.RunFonts();
                        }
                        tblStylePr.RunPropertiesBaseStyle.RunFonts.EastAsia = value;
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.RunFonts == null)
                        {
                            _style.StyleRunProperties.RunFonts = new W.RunFonts();
                        }
                        _style.StyleRunProperties.RunFonts.EastAsia = value;
                    }
                }
                else if(_paragraph != null)
                {
                    if(_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if(_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.RunFonts());
                    W.RunFonts rFonts = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>();
                    rFonts.EastAsia = value;
                }
                else if(_numberingLevel != null)
                {
                    if(_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if(_numberingLevel.NumberingSymbolRunProperties.RunFonts == null)
                        _numberingLevel.NumberingSymbolRunProperties.RunFonts = new W.RunFonts();
                    _numberingLevel.NumberingSymbolRunProperties.RunFonts.EastAsia = value;
                }
                else
                {
                    _fontNameEastAsia = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the font used for Latin text (characters with character codes from
        /// 0 through 127).
        /// </summary>
        public string FontNameAscii
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _fontNameAscii;
                }
                W.RunFonts rFonts = null;
                if (_run?.RunProperties?.RunFonts != null)
                {
                    rFonts = _run.RunProperties.RunFonts;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        rFonts = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.RunFonts;
                    }
                    else
                    {
                        rFonts = _style.StyleRunProperties?.RunFonts;
                    }
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.RunFonts != null)
                {
                    rFonts = _defaultRPr.RunPropertiesBaseStyle.RunFonts;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.RunFonts>() != null)
                {
                    rFonts = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.RunFonts != null)
                {
                    rFonts = _numberingLevel.NumberingSymbolRunProperties.RunFonts;
                }
                if (rFonts?.AsciiTheme != null)
                {
                    return _document.GetThemeFont(rFonts.AsciiTheme);
                }
                return rFonts?.Ascii;
            }
            set
            {
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.RunFonts == null)
                    {
                        _run.RunProperties.RunFonts = new W.RunFonts();
                    }
                    _run.RunProperties.RunFonts.Ascii = value;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.RunFonts == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.RunFonts = new W.RunFonts();
                        }
                        tblStylePr.RunPropertiesBaseStyle.RunFonts.Ascii = value;
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.RunFonts == null)
                        {
                            _style.StyleRunProperties.RunFonts = new W.RunFonts();
                        }
                        _style.StyleRunProperties.RunFonts.Ascii = value;
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.RunFonts());
                    W.RunFonts rFonts = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>();
                    rFonts.Ascii = value;
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.RunFonts == null)
                        _numberingLevel.NumberingSymbolRunProperties.RunFonts = new W.RunFonts();
                    _numberingLevel.NumberingSymbolRunProperties.RunFonts.Ascii = value;
                }
                else
                {
                    _fontNameAscii = value;
                }
            }
        }

        public string FontNameHighAnsi
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _fontNameHAnsi;
                }
                W.RunFonts rFonts = null;
                if (_run?.RunProperties?.RunFonts != null)
                {
                    rFonts = _run.RunProperties.RunFonts;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        rFonts = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.RunFonts;
                    }
                    else
                    {
                        rFonts = _style.StyleRunProperties?.RunFonts;
                    }
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.RunFonts != null)
                {
                    rFonts = _defaultRPr.RunPropertiesBaseStyle.RunFonts;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.RunFonts>() != null)
                {
                    rFonts = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.RunFonts != null)
                {
                    rFonts = _numberingLevel.NumberingSymbolRunProperties.RunFonts;
                }
                if (rFonts?.HighAnsiTheme != null)
                {
                    return _document.GetThemeFont(rFonts.HighAnsiTheme);
                }
                return rFonts?.HighAnsi;
            }
            set
            {
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.RunFonts == null)
                    {
                        _run.RunProperties.RunFonts = new W.RunFonts();
                    }
                    _run.RunProperties.RunFonts.HighAnsi = value;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.RunFonts == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.RunFonts = new W.RunFonts();
                        }
                        tblStylePr.RunPropertiesBaseStyle.RunFonts.HighAnsi = value;
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.RunFonts == null)
                        {
                            _style.StyleRunProperties.RunFonts = new W.RunFonts();
                        }
                        _style.StyleRunProperties.RunFonts.HighAnsi = value;
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.RunFonts());
                    W.RunFonts rFonts = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>();
                    rFonts.HighAnsi = value;
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.RunFonts == null)
                        _numberingLevel.NumberingSymbolRunProperties.RunFonts = new W.RunFonts();
                    _numberingLevel.NumberingSymbolRunProperties.RunFonts.HighAnsi = value;
                }
                else
                {
                    _fontNameHAnsi = value;
                }
            }
        }

        public string FontNameComplexScript
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _fontNameCs;
                }
                W.RunFonts rFonts = null;
                if (_run?.RunProperties?.RunFonts != null)
                {
                    rFonts = _run.RunProperties.RunFonts;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        rFonts = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.RunFonts;
                    }
                    else
                    {
                        rFonts = _style.StyleRunProperties?.RunFonts;
                    }
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.RunFonts != null)
                {
                    rFonts = _defaultRPr.RunPropertiesBaseStyle.RunFonts;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.RunFonts>() != null)
                {
                    rFonts = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.RunFonts != null)
                {
                    rFonts = _numberingLevel.NumberingSymbolRunProperties.RunFonts;
                }
                if (rFonts?.ComplexScriptTheme != null)
                {
                    return _document.GetThemeFont(rFonts.ComplexScriptTheme);
                }
                return rFonts?.ComplexScript;
            }
            set
            {
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.RunFonts == null)
                    {
                        _run.RunProperties.RunFonts = new W.RunFonts();
                    }
                    _run.RunProperties.RunFonts.ComplexScript = value;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.RunFonts == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.RunFonts = new W.RunFonts();
                        }
                        tblStylePr.RunPropertiesBaseStyle.RunFonts.ComplexScript = value;
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.RunFonts == null)
                        {
                            _style.StyleRunProperties.RunFonts = new W.RunFonts();
                        }
                        _style.StyleRunProperties.RunFonts.ComplexScript = value;
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.RunFonts());
                    W.RunFonts rFonts = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>();
                    rFonts.ComplexScript = value;
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.RunFonts == null)
                        _numberingLevel.NumberingSymbolRunProperties.RunFonts = new W.RunFonts();
                    _numberingLevel.NumberingSymbolRunProperties.RunFonts.ComplexScript = value;
                }
                else
                {
                    _fontNameHAnsi = value;
                }
            }
        }

        public EnumValue<FontContentType> FontTypeHint
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _fontTypeHint;
                }
                W.RunFonts rFonts = null;
                if (_run?.RunProperties?.RunFonts != null)
                {
                    rFonts = _run.RunProperties.RunFonts;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        rFonts = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.RunFonts;
                    }
                    else
                    {
                        rFonts = _style.StyleRunProperties?.RunFonts;
                    }
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.RunFonts != null)
                {
                    rFonts = _defaultRPr.RunPropertiesBaseStyle.RunFonts;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.RunFonts>() != null)
                {
                    rFonts = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.RunFonts != null)
                {
                    rFonts = _numberingLevel.NumberingSymbolRunProperties.RunFonts;
                }
                if (rFonts?.Hint == null) return null;
                return rFonts.Hint.Value.Convert<FontContentType>();
            }
            set
            {
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.RunFonts == null)
                    {
                        _run.RunProperties.RunFonts = new W.RunFonts();
                    }
                    _run.RunProperties.RunFonts.Hint = value.Val.Convert<W.FontTypeHintValues>();
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.RunFonts == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.RunFonts = new W.RunFonts();
                        }
                        tblStylePr.RunPropertiesBaseStyle.RunFonts.Hint = value.Val.Convert<W.FontTypeHintValues>();
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.RunFonts == null)
                        {
                            _style.StyleRunProperties.RunFonts = new W.RunFonts();
                        }
                        _style.StyleRunProperties.RunFonts.Hint = value.Val.Convert<W.FontTypeHintValues>();
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.RunFonts());
                    W.RunFonts rFonts = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.RunFonts>();
                    rFonts.Hint = value.Val.Convert<W.FontTypeHintValues>();
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.RunFonts == null)
                        _numberingLevel.NumberingSymbolRunProperties.RunFonts = new W.RunFonts();
                    _numberingLevel.NumberingSymbolRunProperties.RunFonts.Hint = value.Val.Convert<W.FontTypeHintValues>();
                }
                else
                {
                    _fontTypeHint = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets font size specified in points.
        /// </summary>
        public FloatValue FontSize
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _fontSize;
                }
                W.FontSize sz = null;
                if (_run?.RunProperties?.FontSize != null)
                {
                    sz = _run.RunProperties.FontSize;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        sz = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.FontSize;
                    }
                    else
                    {
                        sz = _style.StyleRunProperties?.FontSize;
                    }
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.FontSize != null)
                {
                    sz = _defaultRPr.RunPropertiesBaseStyle.FontSize;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.FontSize>() != null)
                {
                    sz = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.FontSize>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.FontSize != null)
                {
                    sz = _numberingLevel.NumberingSymbolRunProperties.FontSize;
                }
                if (sz == null) return null;
                return sz.Val.Value.ToFloat() / 2;
            }
            set
            {
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.FontSize == null)
                    {
                        _run.RunProperties.FontSize = new W.FontSize();
                    }
                    _run.RunProperties.FontSize.Val = (value * 2).ToString();
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.FontSize == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.FontSize = new W.FontSize();
                        }
                        tblStylePr.RunPropertiesBaseStyle.FontSize.Val = (value * 2).ToString();
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.FontSize == null)
                        {
                            _style.StyleRunProperties.FontSize = new W.FontSize();
                        }
                        _style.StyleRunProperties.FontSize.Val = (value * 2).ToString();
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.FontSize>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.FontSize());
                    W.FontSize sz = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.FontSize>();
                    sz.Val = (value * 2).ToString();
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.FontSize == null)
                        _numberingLevel.NumberingSymbolRunProperties.FontSize = new W.FontSize();
                    _numberingLevel.NumberingSymbolRunProperties.FontSize.Val = (value * 2).ToString();
                }
                else
                {
                    _fontSize = value;
                }
            }
        }

        public FloatValue FontSizeCs
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _fontSizeCs;
                }
                W.FontSizeComplexScript sz = null;
                if (_run?.RunProperties?.FontSizeComplexScript != null)
                {
                    sz = _run.RunProperties.FontSizeComplexScript;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        sz = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.FontSizeComplexScript;
                    }
                    else
                    {
                        sz = _style.StyleRunProperties?.FontSizeComplexScript;
                    }
                }
                else if (_style?.StyleRunProperties?.FontSizeComplexScript != null)
                {
                    sz = _style.StyleRunProperties.FontSizeComplexScript;
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.FontSizeComplexScript != null)
                {
                    sz = _defaultRPr.RunPropertiesBaseStyle.FontSizeComplexScript;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.FontSizeComplexScript>() != null)
                {
                    sz = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.FontSizeComplexScript>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.FontSizeComplexScript != null)
                {
                    sz = _numberingLevel.NumberingSymbolRunProperties.FontSizeComplexScript;
                }
                if (sz == null) return null;
                return sz.Val.Value.ToFloat() / 2;
            }
            set
            {
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.FontSizeComplexScript == null)
                    {
                        _run.RunProperties.FontSizeComplexScript = new W.FontSizeComplexScript();
                    }
                    _run.RunProperties.FontSizeComplexScript.Val = (value * 2).ToString();
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.FontSizeComplexScript == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.FontSizeComplexScript = new W.FontSizeComplexScript();
                        }
                        tblStylePr.RunPropertiesBaseStyle.FontSizeComplexScript.Val = (value * 2).ToString();
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.FontSizeComplexScript == null)
                        {
                            _style.StyleRunProperties.FontSizeComplexScript = new W.FontSizeComplexScript();
                        }
                        _style.StyleRunProperties.FontSizeComplexScript.Val = (value * 2).ToString();
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.FontSizeComplexScript>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.FontSizeComplexScript());
                    W.FontSizeComplexScript sz = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.FontSizeComplexScript>();
                    sz.Val = (value * 2).ToString();
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.FontSizeComplexScript == null)
                        _numberingLevel.NumberingSymbolRunProperties.FontSizeComplexScript = new W.FontSizeComplexScript();
                    _numberingLevel.NumberingSymbolRunProperties.FontSizeComplexScript.Val = (value * 2).ToString();
                }
                else
                {
                    _fontSizeCs = value;
                }
            }
        }
        /// <summary>
        /// Gets or sets bold style.
        /// </summary>
        public BooleanValue Bold
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _bold;
                }
                W.Bold bold = null;
                if (_run?.RunProperties?.Bold != null)
                {
                    bold = _run.RunProperties.Bold;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        bold = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.Bold;
                    }
                    else
                    {
                        bold = _style.StyleRunProperties?.Bold;
                    }
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.Bold != null)
                {
                    bold = _defaultRPr.RunPropertiesBaseStyle.Bold;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.Bold>() != null)
                {
                    bold = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Bold>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.Bold != null)
                {
                    bold = _numberingLevel.NumberingSymbolRunProperties.Bold;
                }
                if (bold == null) return null;
                if (bold.Val == null) return true;
                return bold.Val.Value;
            }
            set
            {
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.Bold == null)
                    {
                        _run.RunProperties.Bold = new W.Bold();
                    }
                    if (value)
                    {
                        _run.RunProperties.Bold.Val = null;
                    }
                    else
                    {
                        _run.RunProperties.Bold.Val = false;
                    }
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.Bold == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.Bold = new W.Bold();
                        }
                        if(value)
                            tblStylePr.RunPropertiesBaseStyle.Bold.Val = null;
                        else
                            tblStylePr.RunPropertiesBaseStyle.Bold.Val = false;
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.Bold == null)
                        {
                            _style.StyleRunProperties.Bold = new W.Bold();
                        }
                        if(value)
                            _style.StyleRunProperties.Bold.Val = null;
                        else
                            _style.StyleRunProperties.Bold.Val = false;
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Bold>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.Bold());
                    W.Bold bold = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Bold>();
                    if (value) bold.Val = null;
                    else bold.Val = false;
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.Bold == null)
                        _numberingLevel.NumberingSymbolRunProperties.Bold = new W.Bold();
                    if (value) _numberingLevel.NumberingSymbolRunProperties.Bold.Val = null;
                    else _numberingLevel.NumberingSymbolRunProperties.Bold.Val = false;
                }
                else
                {
                    _bold = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets italic style.
        /// </summary>
        public BooleanValue Italic
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _italic;
                }
                W.Italic italic = null;
                if (_run?.RunProperties?.Italic != null)
                {
                    italic = _run.RunProperties.Italic;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        italic = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.Italic;
                    }
                    else
                    {
                        italic = _style.StyleRunProperties?.Italic;
                    }
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.Italic != null)
                {
                    italic = _defaultRPr.RunPropertiesBaseStyle.Italic;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.Italic>() != null)
                {
                    italic = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Italic>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.Italic != null)
                {
                    italic = _numberingLevel.NumberingSymbolRunProperties.Italic;
                }
                if (italic == null) return null;
                if (italic.Val == null) return true;
                return italic.Val.Value;
            }
            set
            {
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.Italic == null)
                    {
                        _run.RunProperties.Italic = new W.Italic();
                    }
                    if (value)
                    {
                        _run.RunProperties.Italic.Val = null;
                    }
                    else
                    {
                        _run.RunProperties.Italic.Val = false;
                    }
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.Italic == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.Italic = new W.Italic();
                        }
                        if (value)
                            tblStylePr.RunPropertiesBaseStyle.Italic.Val = null;
                        else
                            tblStylePr.RunPropertiesBaseStyle.Italic.Val = false;
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.Italic == null)
                        {
                            _style.StyleRunProperties.Italic = new W.Italic();
                        }
                        if (value)
                            _style.StyleRunProperties.Italic.Val = null;
                        else
                            _style.StyleRunProperties.Italic.Val = false;
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Italic>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.Italic());
                    W.Italic italic = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Italic>();
                    if (value) italic.Val = null;
                    else italic.Val = false;
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.Italic == null)
                        _numberingLevel.NumberingSymbolRunProperties.Italic = new W.Italic();
                    if (value) _numberingLevel.NumberingSymbolRunProperties.Italic.Val = null;
                    else _numberingLevel.NumberingSymbolRunProperties.Italic.Val = false;
                }
                else
                {
                    _italic = value;
                }
            }
        }

        public EnumValue<SubSuperScript> SubSuperScript
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _subSuperScript;
                }
                W.VerticalTextAlignment vAlign = null;
                if (_run?.RunProperties?.VerticalTextAlignment != null)
                {
                    vAlign = _run.RunProperties.VerticalTextAlignment;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        vAlign = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.VerticalTextAlignment;
                    }
                    else
                    {
                        vAlign = _style.StyleRunProperties?.VerticalTextAlignment;
                    }
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.VerticalTextAlignment != null)
                {
                    vAlign = _defaultRPr.RunPropertiesBaseStyle.VerticalTextAlignment;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.VerticalTextAlignment>() != null)
                {
                    vAlign = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.VerticalTextAlignment>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.VerticalTextAlignment != null)
                {
                    vAlign = _numberingLevel.NumberingSymbolRunProperties.VerticalTextAlignment;
                }
                if (vAlign == null) return null;
                if (vAlign.Val == null || vAlign.Val.Value == W.VerticalPositionValues.Baseline) return Berry.Docx.SubSuperScript.None;
                return vAlign.Val.Value.Convert<SubSuperScript>();
            }
            set
            {
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.VerticalTextAlignment == null)
                    {
                        _run.RunProperties.VerticalTextAlignment = new W.VerticalTextAlignment();
                    }
                    _run.RunProperties.VerticalTextAlignment.Val = value.Val.Convert(W.VerticalPositionValues.Baseline);
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.VerticalTextAlignment == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.VerticalTextAlignment = new W.VerticalTextAlignment();
                        }
                        tblStylePr.RunPropertiesBaseStyle.VerticalTextAlignment.Val = value.Val.Convert(W.VerticalPositionValues.Baseline);
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.VerticalTextAlignment == null)
                        {
                            _style.StyleRunProperties.VerticalTextAlignment = new W.VerticalTextAlignment();
                        }
                        _style.StyleRunProperties.VerticalTextAlignment.Val = value.Val.Convert(W.VerticalPositionValues.Baseline);
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.VerticalTextAlignment>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.VerticalTextAlignment());
                    W.VerticalTextAlignment vAlign = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.VerticalTextAlignment>();
                    vAlign.Val = value.Val.Convert(W.VerticalPositionValues.Baseline);
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.VerticalTextAlignment == null)
                        _numberingLevel.NumberingSymbolRunProperties.VerticalTextAlignment = new W.VerticalTextAlignment();
                    _numberingLevel.NumberingSymbolRunProperties.VerticalTextAlignment.Val = value.Val.Convert(W.VerticalPositionValues.Baseline);
                }
                else
                {
                    _subSuperScript = value;
                }
            }
        }

        public EnumValue<UnderlineStyle> UnderlineStyle
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _underlineStyle;
                }
                W.Underline underline = null;
                if (_run?.RunProperties?.Underline != null)
                {
                    underline = _run.RunProperties.Underline;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        underline = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.Underline;
                    }
                    else
                    {
                        underline = _style.StyleRunProperties?.Underline;
                    }
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.Underline != null)
                {
                    underline = _defaultRPr.RunPropertiesBaseStyle.Underline;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.Underline>() != null)
                {
                    underline = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Underline>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.Underline != null)
                {
                    underline = _numberingLevel.NumberingSymbolRunProperties.Underline;
                }
                if (underline == null) return null;
                if (underline.Val == null) return Berry.Docx.UnderlineStyle.None;
                return underline.Val.Value.Convert<UnderlineStyle>();
            }
            set
            {
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.Underline == null)
                    {
                        _run.RunProperties.Underline = new W.Underline();
                    }
                    _run.RunProperties.Underline.Val = value.Val.Convert<W.UnderlineValues>();
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.Underline == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.Underline = new W.Underline();
                        }
                        tblStylePr.RunPropertiesBaseStyle.Underline.Val = value.Val.Convert<W.UnderlineValues>();
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.Underline == null)
                        {
                            _style.StyleRunProperties.Underline = new W.Underline();
                        }
                        _style.StyleRunProperties.Underline.Val = value.Val.Convert<W.UnderlineValues>();
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Underline>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.Underline());
                    W.Underline underline = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Underline>();
                    underline.Val = value.Val.Convert<W.UnderlineValues>();
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.Underline == null)
                        _numberingLevel.NumberingSymbolRunProperties.Underline = new W.Underline();
                    _numberingLevel.NumberingSymbolRunProperties.Underline.Val = value.Val.Convert<W.UnderlineValues>();
                }
                else
                {
                    _underlineStyle = value;
                }
            }
        }

        public ColorValue TextColor
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _color;
                }
                W.Color color = null;
                if (_run?.RunProperties?.Color != null)
                {
                    color = _run.RunProperties.Color;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        color = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.Color;
                    }
                    else
                    {
                        color = _style.StyleRunProperties?.Color;
                    }
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.Color != null)
                {
                    color = _defaultRPr.RunPropertiesBaseStyle.Color;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.Color>() != null)
                {
                    color = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Color>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.Color != null)
                {
                    color = _numberingLevel.NumberingSymbolRunProperties.Color;
                }
                if (color?.Val != null)
                {
                    return color.Val.Value;
                }
                return null;
            }
            set
            {
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.Color == null)
                    {
                        _run.RunProperties.Color = new W.Color();
                    }
                    _run.RunProperties.Color.Val = value.ToString();
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.Color == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.Color = new W.Color();
                        }
                        tblStylePr.RunPropertiesBaseStyle.Color.Val = value.ToString();
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.Color == null)
                        {
                            _style.StyleRunProperties.Color = new W.Color();
                        }
                        _style.StyleRunProperties.Color.Val = value.ToString();
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Color>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.Color());
                    W.Color color = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Color>();
                    color.Val = value.ToString();
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.Color == null)
                        _numberingLevel.NumberingSymbolRunProperties.Color = new W.Color();
                    _numberingLevel.NumberingSymbolRunProperties.Color.Val = value.ToString();
                }
                else
                {
                    _color = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the percent value of the normal character width that each character shall be scaled.
        /// <para>If the value is 100, then each character shall be displayed at 100% of its normal with.</para>
        /// <para>The value must be between 1 and 600, otherwise an exception will be thrown.</para>
        /// </summary>
        /// <exception cref="InvalidOperationException"/>
        public IntegerValue CharacterScale
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _characterScale;
                }
                W.CharacterScale scale = null;
                if (_run?.RunProperties?.CharacterScale != null)
                {
                    scale = _run.RunProperties.CharacterScale;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        scale = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.CharacterScale;
                    }
                    else
                    {
                        scale = _style.StyleRunProperties?.CharacterScale;
                    }
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.CharacterScale != null)
                {
                    scale = _defaultRPr.RunPropertiesBaseStyle.CharacterScale;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.CharacterScale>() != null)
                {
                    scale = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.CharacterScale>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.CharacterScale != null)
                {
                    scale = _numberingLevel.NumberingSymbolRunProperties.CharacterScale;
                }
                if (scale == null) return null;
                return (int)scale.Val;
            }
            set
            {
                if (value != null && (value < 1 || value > 600))
                {
                    throw new InvalidOperationException("This is not a vaild measurement. The value must be between 1 and 600.");
                }
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.CharacterScale == null)
                    {
                        _run.RunProperties.CharacterScale = new W.CharacterScale();
                    }
                    _run.RunProperties.CharacterScale.Val = value.Val;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.CharacterScale == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.CharacterScale = new W.CharacterScale();
                        }
                        tblStylePr.RunPropertiesBaseStyle.CharacterScale.Val = value.Val;
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.CharacterScale == null)
                        {
                            _style.StyleRunProperties.CharacterScale = new W.CharacterScale();
                        }
                        _style.StyleRunProperties.CharacterScale.Val = value.Val;
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.CharacterScale>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.CharacterScale());
                    W.CharacterScale scale = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.CharacterScale>();
                    scale.Val = value.Val;
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.CharacterScale == null)
                        _numberingLevel.NumberingSymbolRunProperties.CharacterScale = new W.CharacterScale();
                    _numberingLevel.NumberingSymbolRunProperties.CharacterScale.Val = value.Val;
                }
                else
                {
                    _characterScale = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the amount (in points) of character pitch which shall be added or removed after each character.
        /// </summary>
        public FloatValue CharacterSpacing
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _characterSpacing;
                }
                W.Spacing spacing = null;
                if (_run?.RunProperties?.Spacing != null)
                {
                    spacing = _run.RunProperties.Spacing;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        spacing = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.Spacing;
                    }
                    else
                    {
                        spacing = _style.StyleRunProperties?.Spacing;
                    }
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.Spacing != null)
                {
                    spacing = _defaultRPr.RunPropertiesBaseStyle.Spacing;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.Spacing>() != null)
                {
                    spacing = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Spacing>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.Spacing != null)
                {
                    spacing = _numberingLevel.NumberingSymbolRunProperties.Spacing;
                }
                if (spacing == null) return null;
                return spacing.Val / 20.0F;
            }
            set
            {
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.Spacing == null)
                    {
                        _run.RunProperties.Spacing = new W.Spacing();
                    }
                    _run.RunProperties.Spacing.Val = (int)(value * 20);
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.Spacing == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.Spacing = new W.Spacing();
                        }
                        tblStylePr.RunPropertiesBaseStyle.Spacing.Val = Convert.ToInt32(value * 20);
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.Spacing == null)
                        {
                            _style.StyleRunProperties.Spacing = new W.Spacing();
                        }
                        _style.StyleRunProperties.Spacing.Val = Convert.ToInt32(value * 20);
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Spacing>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.Spacing());
                    W.Spacing spacing = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Spacing>();
                    spacing.Val = (int)(value * 20);
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.Spacing == null)
                        _numberingLevel.NumberingSymbolRunProperties.Spacing = new W.Spacing();
                    _numberingLevel.NumberingSymbolRunProperties.Spacing.Val = (int)(value * 20);
                }
                else
                {
                    _characterSpacing = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the amount (in points) by which text shall be raised or lowered in relation to the default baseline location.
        /// </summary>
        public FloatValue Position
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _position;
                }
                W.Position position = null;
                if (_run?.RunProperties?.Position != null)
                {
                    position = _run.RunProperties.Position;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        position = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.Position;
                    }
                    else
                    {
                        position = _style.StyleRunProperties?.Position;
                    }
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.Position != null)
                {
                    position = _defaultRPr.RunPropertiesBaseStyle.Position;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.Position>() != null)
                {
                    position = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Position>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.Position != null)
                {
                    position = _numberingLevel.NumberingSymbolRunProperties.Position;
                }
                if (position == null) return null;
                return position.Val.ToString().ToFloat() / 2;
            }
            set
            {
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.Position == null)
                    {
                        _run.RunProperties.Position = new W.Position();
                    }
                    _run.RunProperties.Position.Val = Math.Round(value * 2).ToString();
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.Position == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.Position = new W.Position();
                        }
                        tblStylePr.RunPropertiesBaseStyle.Position.Val = Math.Round(value * 2).ToString();
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.Position == null)
                        {
                            _style.StyleRunProperties.Position = new W.Position();
                        }
                        _style.StyleRunProperties.Position.Val = Math.Round(value * 2).ToString();
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Position>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.Position());
                    W.Position pos = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Position>();
                    pos.Val = Math.Round(value * 2).ToString();
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.Position == null)
                        _numberingLevel.NumberingSymbolRunProperties.Position = new W.Position();
                    _numberingLevel.NumberingSymbolRunProperties.Position.Val = Math.Round(value * 2).ToString();
                }
                else
                {
                    _position = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the text is hidden.
        /// </summary>
        public BooleanValue IsHidden
        {
            get
            {
                if (_run == null && _style == null && _defaultRPr == null && _paragraph == null && _numberingLevel == null)
                {
                    return _isHidden;
                }
                W.Vanish vanish = null;
                if (_run?.RunProperties?.Vanish != null)
                {
                    vanish = _run.RunProperties.Vanish;
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        vanish = _style.Elements<W.TableStyleProperties>()
                                .Where(t => t.Type == _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>()).FirstOrDefault()
                                ?.RunPropertiesBaseStyle?.Vanish;
                    }
                    else
                    {
                        vanish = _style.StyleRunProperties?.Vanish;
                    }
                }
                else if (_defaultRPr?.RunPropertiesBaseStyle?.Vanish != null)
                {
                    vanish = _defaultRPr.RunPropertiesBaseStyle.Vanish;
                }
                else if (_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.Vanish>() != null)
                {
                    vanish = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Vanish>();
                }
                else if (_numberingLevel?.NumberingSymbolRunProperties?.Vanish != null)
                {
                    vanish = _numberingLevel.NumberingSymbolRunProperties.Vanish;
                }
                if (vanish == null) return null;
                if (vanish.Val == null) return true;
                return vanish.Val.Value;
            }
            set
            {
                if (_run != null)
                {
                    if (_run.RunProperties == null)
                    {
                        _run.RunProperties = new W.RunProperties();
                    }
                    if (_run.RunProperties.Vanish == null)
                    {
                        _run.RunProperties.Vanish = new W.Vanish();
                    }
                    if (value)
                    {
                        _run.RunProperties.Vanish.Val = null;
                    }
                    else
                    {
                        _run.RunProperties.Vanish.Val = false;
                    }
                }
                else if (_style != null)
                {
                    if (_tableStyleRegion != null && _tableStyleRegion != TableRegionType.WholeTable)
                    {
                        W.TableStyleOverrideValues type = _tableStyleRegion.Val.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                        {
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        }
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.RunPropertiesBaseStyle == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle = new W.RunPropertiesBaseStyle();
                        }
                        if (tblStylePr.RunPropertiesBaseStyle.Vanish == null)
                        {
                            tblStylePr.RunPropertiesBaseStyle.Vanish = new W.Vanish();
                        }
                        if (value)
                            tblStylePr.RunPropertiesBaseStyle.Vanish.Val = null;
                        else
                            tblStylePr.RunPropertiesBaseStyle.Vanish.Val = false;
                    }
                    else
                    {
                        if (_style.StyleRunProperties == null)
                        {
                            _style.StyleRunProperties = new W.StyleRunProperties();
                        }
                        if (_style.StyleRunProperties.Vanish == null)
                        {
                            _style.StyleRunProperties.Vanish = new W.Vanish();
                        }
                        if (value)
                            _style.StyleRunProperties.Vanish.Val = null;
                        else
                            _style.StyleRunProperties.Vanish.Val = false;
                    }
                }
                else if (_paragraph != null)
                {
                    if (_paragraph.ParagraphProperties == null)
                        _paragraph.ParagraphProperties = new W.ParagraphProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties = new W.ParagraphMarkRunProperties();
                    if (_paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Vanish>() == null)
                        _paragraph.ParagraphProperties.ParagraphMarkRunProperties.AddChild(new W.Vanish());
                    W.Vanish vanish = _paragraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<W.Vanish>();
                    if (value) vanish.Val = null;
                    else vanish.Val = false;
                }
                else if (_numberingLevel != null)
                {
                    if (_numberingLevel.NumberingSymbolRunProperties == null)
                        _numberingLevel.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties();
                    if (_numberingLevel.NumberingSymbolRunProperties.Vanish == null)
                        _numberingLevel.NumberingSymbolRunProperties.Vanish = new W.Vanish();
                    if (value) _numberingLevel.NumberingSymbolRunProperties.Vanish.Val = null;
                    else _numberingLevel.NumberingSymbolRunProperties.Vanish.Val = false;
                }
                else
                {
                    _isHidden = value;
                }
            }
        }

        #endregion

        #region Public Methods
        /// <summary>
        /// Clears all character formats.
        /// </summary>
        public void ClearFormatting()
        {
            if (_run?.RunProperties != null)
            {
                _run.RunProperties = null;
            }
            else if (_style?.StyleRunProperties != null)
            {
                _style.StyleRunProperties.RemoveAllChildren();
            }
            else if(_paragraph?.ParagraphProperties?.ParagraphMarkRunProperties != null)
            {
                _paragraph.ParagraphProperties.ParagraphMarkRunProperties.RemoveAllChildren();
            }
            else if(_numberingLevel?.NumberingSymbolRunProperties != null)
            {
                _numberingLevel.NumberingSymbolRunProperties.RemoveAllChildren();
            }
        }

        public static RunPropertiesHolder GetRunStyleFormatRecursively(Document doc, W.Style style)
        {
            RunPropertiesHolder format = new RunPropertiesHolder();
            RunPropertiesHolder baseFmt = new RunPropertiesHolder();
            W.Style baseStyle = style.GetBaseStyle(doc);
            if(baseStyle != null)
            {
                baseFmt = GetRunStyleFormatRecursively(doc, baseStyle);
            }
            RunPropertiesHolder directFmt = new RunPropertiesHolder(doc.Package, style);

            format.FontNameAscii = directFmt.FontNameAscii ?? baseFmt.FontNameAscii;
            format.FontNameEastAsia = directFmt.FontNameEastAsia ?? baseFmt.FontNameEastAsia;
            format.FontNameHighAnsi = directFmt.FontNameHighAnsi ?? baseFmt.FontNameHighAnsi;
            format.FontNameComplexScript = directFmt.FontNameComplexScript ?? baseFmt.FontNameComplexScript;
            format.FontTypeHint = directFmt.FontTypeHint ?? baseFmt.FontTypeHint;

            format.FontSize = directFmt.FontSize ?? baseFmt.FontSize;
            format.FontSizeCs = directFmt.FontSizeCs ?? baseFmt.FontSizeCs;
            format.Bold = directFmt.Bold ?? baseFmt.Bold;
            format.Italic = directFmt.Italic ?? baseFmt.Italic;
            format.SubSuperScript = directFmt.SubSuperScript ?? baseFmt.SubSuperScript;
            format.UnderlineStyle = directFmt.UnderlineStyle ?? baseFmt.UnderlineStyle;
            format.TextColor = directFmt.TextColor ?? baseFmt.TextColor;
            format.CharacterScale = directFmt.CharacterScale ?? baseFmt.CharacterScale;
            format.CharacterSpacing = directFmt.CharacterSpacing ?? baseFmt.CharacterSpacing;
            format.Position = directFmt.Position ?? baseFmt.Position;

            return format;
        }
        #endregion
    }
}
