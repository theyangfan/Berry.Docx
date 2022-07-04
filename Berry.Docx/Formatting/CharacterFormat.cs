using System;
using System.Linq;
using System.Drawing;
using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the character format.
    /// </summary>
    public class CharacterFormat
    {
        #region Private Members

        private Document _doc;

        #region TextRange
        private W.Run _ownerRun;
        private RunPropertiesHolder _directRHld;
        #endregion

        #region Paragraph
        private W.Paragraph _ownerParagraph;
        private readonly RunPropertiesHolder _markRHld;
        #endregion

        #region Style
        private W.Style _ownerStyle;
        private RunPropertiesHolder _directSHld;
        #endregion

        #region Numbering
        private W.Level _numberingLevel;
        private RunPropertiesHolder _numRHld;
        #endregion

        #region Formats
        private string _fontCN = "宋体";
        private string _fontEN = "Times New Roman";
        private float _fontSize = 10.5F;
        private FontContentType _fontTypeHint = FontContentType.Default;
        private float _fontSizeCs = 10.5F;
        private bool _bold = false;
        private bool _italic = false;
        private SubSuperScript _subSuperScript = SubSuperScript.None;
        private UnderlineStyle _underlineStyle = UnderlineStyle.None;
        private ColorValue _textColor = ColorValue.Auto;
        private int _characterScale = 100;
        private float _characterSpacing = 0;
        private float _position = 0;
        private Border _border;
        #endregion

        #endregion

        #region Constructors
        internal CharacterFormat() { }

        /// <summary>
        /// Represent the character format of a TextRange.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ownerRun"></param>
        internal CharacterFormat(Document doc, W.Run ownerRun)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _directRHld = new RunPropertiesHolder(doc.Package, ownerRun);
            _border = new Border(doc, ownerRun);
        }

        /// <summary>
        /// Represent the character format of a Paragraph.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ownerParagraph"></param>
        /// <param name="paragraphMark"></param>
        internal CharacterFormat(Document doc, W.Paragraph ownerParagraph)
        {
            _doc = doc;
            _ownerParagraph = ownerParagraph;
            _markRHld = new RunPropertiesHolder(doc.Package, ownerParagraph);
            _border = new Border(doc, ownerParagraph);
        }

        /// <summary>
        /// Represent the character format of a Style.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ownerStyle"></param>
        internal CharacterFormat(Document doc, W.Style ownerStyle)
        {
            _doc = doc;
            _ownerStyle = ownerStyle;
            _directSHld = new RunPropertiesHolder(doc.Package, ownerStyle);
            _border = new Border(doc, ownerStyle);
        }

        internal CharacterFormat(Document doc, W.Level numberingLevel)
        {
            _doc = doc;
            _numberingLevel = numberingLevel;
            _numRHld = new RunPropertiesHolder(doc.Package, numberingLevel);
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
                if(_ownerRun != null)
                {
                    // direct formatting
                    if (_directRHld.FontNameEastAsia != null)
                    {
                        return _directRHld.FontNameEastAsia;
                    }
                    // character style
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        RunPropertiesHolder rStyle = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerRun.GetStyle(_doc));
                        if (rStyle.FontNameEastAsia != null)
                            return rStyle.FontNameEastAsia;
                    }
                    // paragraph style
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively
                            (_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc));
                        if(paragraph.FontNameEastAsia != null)
                            return paragraph.FontNameEastAsia;
                    }
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.FontNameEastAsia;
                }
                else if(_ownerParagraph != null)
                {
                    // paragraph mark
                    if (_markRHld.FontNameEastAsia != null)
                    {
                        return _markRHld.FontNameEastAsia;
                    }
                    // paragraph style
                    RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (paragraph.FontNameEastAsia != null)
                        return paragraph.FontNameEastAsia;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.FontNameEastAsia;
                }
                else if(_ownerStyle != null)
                {
                    // direct formatting
                    if (_directSHld.FontNameEastAsia != null)
                    {
                        return _directSHld.FontNameEastAsia;
                    }
                    // character & paragraph styles
                    RunPropertiesHolder style = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerStyle);
                    if(style.FontNameEastAsia != null)
                        return style.FontNameEastAsia;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.FontNameEastAsia;
                }
                else if(_numberingLevel != null)
                {
                    return _numRHld.FontNameEastAsia ?? _doc.DefaultFormat.CharacterFormat.FontNameEastAsia;
                }
                else
                {
                    return _fontCN;
                }
            }
            set
            {
                if (_ownerRun != null)
                {
                    _directRHld.FontNameEastAsia = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.FontNameEastAsia = value;
                }
                else if(_ownerParagraph != null)
                {
                    _markRHld.FontNameEastAsia = value;
                }
                else if(_numberingLevel != null)
                {
                    _numRHld.FontNameEastAsia = value;
                }
                else
                {
                    _fontCN = value;
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
                if (_ownerRun != null)
                {
                    // direct formatting
                    if (_directRHld.FontNameAscii != null)
                    {
                        return _directRHld.FontNameAscii;
                    }
                    // character style
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        RunPropertiesHolder rStyle = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerRun.GetStyle(_doc));
                        if (rStyle.FontNameAscii != null)
                            return rStyle.FontNameAscii;
                    }
                    // paragraph style
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively
                            (_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc));
                        if (paragraph.FontNameAscii != null)
                            return paragraph.FontNameAscii;
                    }
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.FontNameAscii;
                }
                else if (_ownerParagraph != null)
                {
                    // paragraph mark
                    if (_markRHld.FontNameAscii != null)
                    {
                        return _markRHld.FontNameAscii;
                    }
                    // paragraph style
                    RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (paragraph.FontNameAscii != null)
                        return paragraph.FontNameAscii;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.FontNameAscii;
                }
                else if (_ownerStyle != null)
                {
                    // direct formatting
                    if (_directSHld.FontNameAscii != null)
                    {
                        return _directSHld.FontNameAscii;
                    }
                    // character & paragraph style
                    RunPropertiesHolder style = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerStyle);
                    if (style.FontNameAscii != null)
                        return style.FontNameAscii;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.FontNameAscii;
                }
                else if (_numberingLevel != null)
                {
                    return _numRHld.FontNameAscii ?? _doc.DefaultFormat.CharacterFormat.FontNameAscii;
                }
                else
                {
                    return _fontEN;
                }
            }
            set
            {
                if (_ownerRun != null)
                {
                    _directRHld.FontNameAscii = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.FontNameAscii = value;
                }
                else if (_ownerParagraph != null)
                {
                    _markRHld.FontNameAscii = value;
                }
                else if (_numberingLevel != null)
                {
                    _numRHld.FontNameAscii = value;
                }
                else
                {
                    _fontEN = value;
                }
            }
        }

        public FontContentType FontTypeHint
        {
            get
            {
                if(_ownerRun != null)
                {
                    return _directRHld.FontTypeHint ?? FontContentType.Default;
                }
                else if(_ownerParagraph != null)
                {
                    return _markRHld.FontTypeHint ?? FontContentType.Default;
                }
                else if(_ownerStyle != null)
                {
                    return _directSHld.FontTypeHint ?? FontContentType.Default;
                }
                else if (_numberingLevel != null)
                {
                    return _numRHld.FontTypeHint ?? _doc.DefaultFormat.CharacterFormat.FontTypeHint;
                }
                return _fontTypeHint;
            }
            set
            {
                if (_ownerRun != null)
                {
                    _directRHld.FontTypeHint = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.FontTypeHint = value;
                }
                else if (_ownerParagraph != null)
                {
                    _markRHld.FontTypeHint = value;
                }
                else if (_numberingLevel != null)
                {
                    _numRHld.FontTypeHint = value;
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
        public float FontSize
        {
            get
            {
                if (_ownerRun != null)
                {
                    // direct formatting
                    if (_directRHld.FontSize != null)
                    {
                        return _directRHld.FontSize;
                    }
                    // character style
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        RunPropertiesHolder rStyle = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerRun.GetStyle(_doc));
                        if (rStyle.FontSize != null)
                            return rStyle.FontSize;
                    }
                    // paragraph style
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively
                            (_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc));
                        if (paragraph.FontSize != null)
                            return paragraph.FontSize;
                    }
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.FontSize;
                }
                else if (_ownerParagraph != null)
                {
                    // paragraph mark
                    if (_markRHld.FontSize != null)
                    {
                        return _markRHld.FontSize;
                    }
                    // paragraph style
                    RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (paragraph.FontSize != null)
                        return paragraph.FontSize;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.FontSize;
                }
                else if (_ownerStyle != null)
                {
                    // direct formatting
                    if (_directSHld.FontSize != null)
                    {
                        return _directSHld.FontSize;
                    }
                    // character & paragraph style
                    RunPropertiesHolder style = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerStyle);
                    if (style.FontSize != null)
                        return style.FontSize;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.FontSize;
                }
                else if (_numberingLevel != null)
                {
                    return _numRHld.FontSize ?? _doc.DefaultFormat.CharacterFormat.FontSize;
                }
                else
                {
                    return _fontSize;
                }
            }
            set
            {
                if(_ownerRun != null)
                {
                    _directRHld.FontSize = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.FontSize = value;
                }
                else if (_ownerParagraph != null)
                {
                    _markRHld.FontSize = value;
                }
                else if (_numberingLevel != null)
                {
                    _numRHld.FontSize = value;
                }
                else
                {
                    _fontSize = value;
                }
            }
        }
        /// <summary>
        /// Gets or sets font size in chinese.
        /// </summary>
        public string FontSizeCN
        {
            get
            {
                Dictionary<float, string> sizeList = new Dictionary<float, string> { 
                    { 6.5F, "小六" }, { 7.5F, "六号" }, { 9, "小五" }, { 10.5F, "五号" }, { 12, "小四" }, 
                    { 14, "四号" }, { 15, "小三" }, { 16, "三号" }, { 18, "小二" }, { 22, "二号" }, 
                    { 24, "小一" }, { 26, "一号" }, { 36, "小初" }, { 42, "初号" }
                };
                float size = FontSize;
                return sizeList.ContainsKey(size) ? sizeList[size] : size.ToString();
            }
            set
            {
                Dictionary<string, float> sizeList = new Dictionary<string, float> {
                    { "小六", 6.5F }, { "六号", 7.5F }, { "小五", 9 }, { "五号", 10.5F }, { "小四", 12 },
                    { "四号", 14 }, { "小三", 15 }, { "三号", 16 }, { "小二", 18 }, { "二号", 22 },
                    { "小一", 24 }, { "一号", 26 }, { "小初", 36 }, { "初号", 42 }
                };
                if (sizeList.ContainsKey(value))
                    FontSize = sizeList[value];
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public float FontSizeCs
        {
            get
            {
                if (_ownerRun != null)
                {
                    // direct formatting
                    if (_directRHld.FontSizeCs != null)
                    {
                        return _directRHld.FontSizeCs;
                    }
                    // character style
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        RunPropertiesHolder rStyle = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerRun.GetStyle(_doc));
                        if (rStyle.FontSizeCs != null)
                            return rStyle.FontSizeCs;
                    }
                    // paragraph style
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively
                            (_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc));
                        if (paragraph.FontSizeCs != null)
                            return paragraph.FontSizeCs;
                    }
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.FontSizeCs;
                }
                else if (_ownerParagraph != null)
                {
                    // paragraph mark
                    if (_markRHld.FontSizeCs != null)
                    {
                        return _markRHld.FontSizeCs;
                    }
                    // paragraph style
                    RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (paragraph.FontSizeCs != null)
                        return paragraph.FontSizeCs;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.FontSizeCs;
                }
                else if (_ownerStyle != null)
                {
                    // direct formatting
                    if (_directSHld.FontSizeCs != null)
                    {
                        return _directSHld.FontSizeCs;
                    }
                    // character & paragraph style
                    RunPropertiesHolder style = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerStyle);
                    if (style.FontSizeCs != null)
                        return style.FontSizeCs;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.FontSizeCs;
                }
                else if (_numberingLevel != null)
                {
                    return _numRHld.FontSizeCs ?? _doc.DefaultFormat.CharacterFormat.FontSizeCs;
                }
                else
                {
                    return _fontSizeCs;
                }
            }
            set
            {
                if (_ownerRun != null)
                {
                    _directRHld.FontSizeCs = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.FontSizeCs = value;
                }
                else if (_ownerParagraph != null)
                {
                    _markRHld.FontSizeCs = value;
                }
                else if (_numberingLevel != null)
                {
                    _numRHld.FontSizeCs = value;
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
        public bool Bold
        {
            get
            {
                if (_ownerRun != null)
                {
                    // direct formatting
                    if (_directRHld.Bold != null)
                    {
                        return _directRHld.Bold;
                    }
                    if (_doc.DefaultFormat.CharacterFormat.Bold)
                        return true;
                    // character & paragraph
                    BooleanValue cVal = null;
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        RunPropertiesHolder rStyle = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerRun.GetStyle(_doc));
                        cVal = rStyle.Bold;
                    }
                    BooleanValue pVal = null;
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively
                            (_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc));
                        pVal = paragraph.Bold;
                    }
                    if (cVal != null && pVal != null)
                        return cVal == pVal ? false : true;
                    else if (cVal != null)
                        return cVal;
                    else if (pVal != null)
                        return pVal;
                    else 
                        return false;
                }
                else if (_ownerParagraph != null)
                {
                    // paragraph mark
                    if (_markRHld.Bold != null)
                    {
                        return _markRHld.Bold;
                    }
                    // paragraph style
                    RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (paragraph.Bold != null)
                        return paragraph.Bold;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.Bold;
                }
                else if (_ownerStyle != null)
                {
                    // direct formatting
                    if (_directSHld.Bold != null)
                    {
                        return _directSHld.Bold;
                    }
                    // character & paragraph style
                    RunPropertiesHolder style = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerStyle);
                    if (style.Bold != null)
                        return style.Bold;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.Bold;
                }
                else if (_numberingLevel != null)
                {
                    return _numRHld.Bold ?? _doc.DefaultFormat.CharacterFormat.Bold;
                }
                else
                {
                    return _bold;
                }
            }
            set
            {
                if(_ownerRun != null)
                {
                    _directRHld.Bold = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.Bold = value;
                }
                else if (_ownerParagraph != null)
                {
                    _markRHld.Bold = value;
                }
                else if (_numberingLevel != null)
                {
                    _numRHld.Bold = value;
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
        public bool Italic
        {
            get
            {
                if (_ownerRun != null)
                {
                    // direct formatting
                    if (_directRHld.Italic != null)
                    {
                        return _directRHld.Italic;
                    }
                    if (_doc.DefaultFormat.CharacterFormat.Italic)
                        return true;
                    // character & paragraph
                    BooleanValue cVal = null;
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        RunPropertiesHolder rStyle = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerRun.GetStyle(_doc));
                        cVal = rStyle.Italic;
                    }
                    BooleanValue pVal = null;
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively
                            (_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc));
                        pVal = paragraph.Italic;
                    }
                    if (cVal != null && pVal != null)
                        return cVal == pVal ? false : true;
                    else if (cVal != null)
                        return cVal;
                    else if (pVal != null)
                        return pVal;
                    else
                        return false;
                }
                else if (_ownerParagraph != null)
                {
                    // paragraph mark
                    if (_markRHld.Italic != null)
                    {
                        return _markRHld.Italic;
                    }
                    // paragraph style
                    RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (paragraph.Italic != null)
                        return paragraph.Italic;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.Italic;
                }
                else if (_ownerStyle != null)
                {
                    // direct formatting
                    if (_directSHld.Italic != null)
                    {
                        return _directSHld.Italic;
                    }
                    // character & paragraph style
                    RunPropertiesHolder style = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerStyle);
                    if (style.Italic != null)
                        return style.Italic;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.Italic;
                }
                else if (_numberingLevel != null)
                {
                    return _numRHld.Italic ?? _doc.DefaultFormat.CharacterFormat.Italic;
                }
                else
                {
                    return _italic;
                }
            }
            set
            {
                if (_ownerRun != null)
                {
                    _directRHld.Italic = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.Italic = value;
                }
                else if (_ownerParagraph != null)
                {
                    _markRHld.Italic = value;
                }
                else if (_numberingLevel != null)
                {
                    _numRHld.Italic = value;
                }
                else
                {
                    _italic = value;
                }
            }
        }

        public SubSuperScript SubSuperScript
        {
            get
            {
                if (_ownerRun != null)
                {
                    // direct formatting
                    if (_directRHld.SubSuperScript != null)
                    {
                        return _directRHld.SubSuperScript;
                    }
                    // character style
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        RunPropertiesHolder rStyle = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerRun.GetStyle(_doc));
                        if (rStyle.SubSuperScript != null)
                            return rStyle.SubSuperScript;
                    }
                    // paragraph style
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively
                            (_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc));
                        if (paragraph.SubSuperScript != null)
                            return paragraph.SubSuperScript;
                    }
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.SubSuperScript;
                }
                else if (_ownerParagraph != null)
                {
                    // paragraph mark
                    if (_markRHld.SubSuperScript != null)
                    {
                        return _markRHld.SubSuperScript;
                    }
                    // paragraph style
                    RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (paragraph.SubSuperScript != null)
                        return paragraph.SubSuperScript;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.SubSuperScript;
                }
                else if (_ownerStyle != null)
                {
                    // direct formatting
                    if (_directSHld.SubSuperScript != null)
                    {
                        return _directSHld.SubSuperScript;
                    }
                    // character & paragraph style
                    RunPropertiesHolder style = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerStyle);
                    if (style.SubSuperScript != null)
                        return style.SubSuperScript;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.SubSuperScript;
                }
                else if (_numberingLevel != null)
                {
                    return _numRHld.SubSuperScript ?? _doc.DefaultFormat.CharacterFormat.SubSuperScript;
                }
                else
                {
                    return _subSuperScript;
                }
            }
            set
            {
                if (_ownerRun != null)
                {
                    _directRHld.SubSuperScript = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.SubSuperScript = value;
                }
                else if (_ownerParagraph != null)
                {
                    _markRHld.SubSuperScript = value;
                }
                else if (_numberingLevel != null)
                {
                    _numRHld.SubSuperScript = value;
                }
                else
                {
                    _subSuperScript = value;
                }
            }
        }

        public UnderlineStyle UnderlineStyle
        {
            get
            {
                if (_ownerRun != null)
                {
                    // direct formatting
                    if (_directRHld.UnderlineStyle != null)
                    {
                        return _directRHld.UnderlineStyle;
                    }
                    // character style
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        RunPropertiesHolder rStyle = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerRun.GetStyle(_doc));
                        if (rStyle.UnderlineStyle != null)
                            return rStyle.UnderlineStyle;
                    }
                    // paragraph style
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively
                            (_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc));
                        if (paragraph.UnderlineStyle != null)
                            return paragraph.UnderlineStyle;
                    }
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.UnderlineStyle;
                }
                else if (_ownerParagraph != null)
                {
                    // paragraph mark
                    if (_markRHld.UnderlineStyle != null)
                    {
                        return _markRHld.UnderlineStyle;
                    }
                    // paragraph style
                    RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (paragraph.UnderlineStyle != null)
                        return paragraph.UnderlineStyle;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.UnderlineStyle;
                }
                else if (_ownerStyle != null)
                {
                    // direct formatting
                    if (_directSHld.UnderlineStyle != null)
                    {
                        return _directSHld.UnderlineStyle;
                    }
                    // character & paragraph style
                    RunPropertiesHolder style = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerStyle);
                    if (style.UnderlineStyle != null)
                        return style.UnderlineStyle;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.UnderlineStyle;
                }
                else if (_numberingLevel != null)
                {
                    return _numRHld.UnderlineStyle ?? _doc.DefaultFormat.CharacterFormat.UnderlineStyle;
                }
                else
                {
                    return _underlineStyle;
                }
            }
            set
            {
                if (_ownerRun != null)
                {
                    _directRHld.UnderlineStyle = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.UnderlineStyle = value;
                }
                else if (_ownerParagraph != null)
                {
                    _markRHld.UnderlineStyle = value;
                }
                else if (_numberingLevel != null)
                {
                    _numRHld.UnderlineStyle = value;
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
                if (_ownerRun != null)
                {
                    // direct formatting
                    if (_directRHld.TextColor != null)
                    {
                        return _directRHld.TextColor;
                    }
                    // character style
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        RunPropertiesHolder rStyle = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerRun.GetStyle(_doc));
                        if (rStyle.TextColor != null)
                            return rStyle.TextColor;
                    }
                    // paragraph style
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively
                            (_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc));
                        if (paragraph.TextColor != null)
                            return paragraph.TextColor;
                    }
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.TextColor;
                }
                else if (_ownerParagraph != null)
                {
                    // paragraph mark
                    if (_markRHld.TextColor != null)
                    {
                        return _markRHld.TextColor;
                    }
                    // paragraph style
                    RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (paragraph.TextColor != null)
                        return paragraph.TextColor;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.TextColor;
                }
                else if (_ownerStyle != null)
                {
                    // direct formatting
                    if (_directSHld.TextColor != null)
                    {
                        return _directSHld.TextColor;
                    }
                    // character & paragraph style
                    RunPropertiesHolder style = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerStyle);
                    if (style.TextColor != null)
                        return style.TextColor;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.TextColor;
                }
                else if (_numberingLevel != null)
                {
                    return _numRHld.TextColor ?? _doc.DefaultFormat.CharacterFormat.TextColor;
                }
                else
                {
                    return _textColor;
                }
            }
            set
            {
                if (_ownerRun != null)
                {
                    _directRHld.TextColor = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.TextColor = value;
                }
                else if (_ownerParagraph != null)
                {
                    _markRHld.TextColor = value;
                }
                else if (_numberingLevel != null)
                {
                    _numRHld.TextColor = value;
                }
                else
                {
                    _textColor = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the percent value of the normal character width that each character shall be scaled.
        /// <para>If the value is 100, then each character shall be displayed at 100% of its normal with.</para>
        /// <para>The value must be between 1 and 600, otherwise an exception will be thrown.</para>
        /// </summary>
        /// <exception cref="InvalidOperationException"/>
        public int CharacterScale
        {
            get
            {
                if (_ownerRun != null)
                {
                    // direct formatting
                    if (_directRHld.CharacterScale != null)
                    {
                        return _directRHld.CharacterScale;
                    }
                    // character style
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        RunPropertiesHolder rStyle = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerRun.GetStyle(_doc));
                        if (rStyle.CharacterScale != null)
                            return rStyle.CharacterScale;
                    }
                    // paragraph style
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively
                            (_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc));
                        if (paragraph.CharacterScale != null)
                            return paragraph.CharacterScale;
                    }
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.CharacterScale;
                }
                else if (_ownerParagraph != null)
                {
                    // paragraph mark
                    if (_markRHld.CharacterScale != null)
                    {
                        return _markRHld.CharacterScale;
                    }
                    // paragraph style
                    RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (paragraph.CharacterScale != null)
                        return paragraph.CharacterScale;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.CharacterScale;
                }
                else if (_ownerStyle != null)
                {
                    // direct formatting
                    if (_directSHld.CharacterScale != null)
                    {
                        return _directSHld.CharacterScale;
                    }
                    // character & paragraph style
                    RunPropertiesHolder style = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerStyle);
                    if (style.CharacterScale != null)
                        return style.CharacterScale;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.CharacterScale;
                }
                else if (_numberingLevel != null)
                {
                    return _numRHld.CharacterScale ?? _doc.DefaultFormat.CharacterFormat.CharacterScale;
                }
                else
                {
                    return _characterScale;
                }
            }
            set
            {
                if (_ownerRun != null)
                {
                    _directRHld.CharacterScale = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.CharacterScale = value;
                }
                else if (_ownerParagraph != null)
                {
                    _markRHld.CharacterScale = value;
                }
                else if (_numberingLevel != null)
                {
                    _numRHld.CharacterScale = value;
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
        public float CharacterSpacing
        {
            get
            {
                if (_ownerRun != null)
                {
                    // direct formatting
                    if (_directRHld.CharacterSpacing != null)
                    {
                        return _directRHld.CharacterSpacing;
                    }
                    // character style
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        RunPropertiesHolder rStyle = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerRun.GetStyle(_doc));
                        if (rStyle.CharacterSpacing != null)
                            return rStyle.CharacterSpacing;
                    }
                    // paragraph style
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively
                            (_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc));
                        if (paragraph.CharacterSpacing != null)
                            return paragraph.CharacterSpacing;
                    }
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.CharacterSpacing;
                }
                else if (_ownerParagraph != null)
                {
                    // paragraph mark
                    if (_markRHld.CharacterSpacing != null)
                    {
                        return _markRHld.CharacterSpacing;
                    }
                    // paragraph style
                    RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                    if (paragraph.CharacterSpacing != null)
                        return paragraph.CharacterSpacing;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.CharacterSpacing;
                }
                else if (_ownerStyle != null)
                {
                    // direct formatting
                    if (_directSHld.CharacterSpacing != null)
                    {
                        return _directSHld.CharacterSpacing;
                    }
                    // character & paragraph style
                    RunPropertiesHolder style = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerStyle);
                    if (style.CharacterSpacing != null)
                        return style.CharacterSpacing;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.CharacterSpacing;
                }
                else if (_numberingLevel != null)
                {
                    return _numRHld.CharacterSpacing ?? _doc.DefaultFormat.CharacterFormat.CharacterSpacing;
                }
                else
                {
                    return _characterSpacing;
                }
            }
            set
            {
                if (_ownerRun != null)
                {
                    _directRHld.CharacterSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.CharacterSpacing = value;
                }
                else if (_ownerParagraph != null)
                {
                    _markRHld.CharacterSpacing = value;
                }
                else if (_numberingLevel != null)
                {
                    _numRHld.CharacterSpacing = value;
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
        public float Position
        {
            get
            {
                if (_ownerRun != null)
                {
                    // direct formatting
                    if (_directRHld.Position != null)
                    {
                        return _directRHld.Position;
                    }
                    // character style
                    if (_ownerRun?.RunProperties?.RunStyle != null)
                    {
                        RunPropertiesHolder rStyle = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc, _ownerRun.GetStyle(_doc));
                        if (rStyle.Position != null)
                            return rStyle.Position;
                    }
                    // paragraph style
                    if (_ownerRun.Ancestors<W.Paragraph>().Any())
                    {
                        RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively
                            (_doc, _ownerRun.Ancestors<W.Paragraph>().First().GetStyle(_doc));
                        if (paragraph.Position != null)
                            return paragraph.Position;
                    }
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.Position;
                }
                else if (_ownerParagraph != null)
                {
                    // paragraph mark
                    if (_markRHld.Position != null)
                    {
                        return _markRHld.Position;
                    }
                    // paragraph style
                    RunPropertiesHolder paragraph = RunPropertiesHolder.GetRunStyleFormatRecursively
                        (_doc, _ownerParagraph.GetStyle(_doc));
                    if (paragraph.Position != null)
                        return paragraph.Position;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.Position;
                }
                else if (_ownerStyle != null)
                {
                    // direct formatting
                    if (_directSHld.Position != null)
                    {
                        return _directSHld.Position;
                    }
                    // character & paragraph style
                    RunPropertiesHolder style = RunPropertiesHolder.GetRunStyleFormatRecursively(_doc,_ownerStyle);
                    if (style.Position != null)
                        return style.Position;
                    // document defaults
                    return _doc.DefaultFormat.CharacterFormat.Position;
                }
                else if (_numberingLevel != null)
                {
                    return _numRHld.Position ?? _doc.DefaultFormat.CharacterFormat.Position;
                }
                else
                {
                    return _position;
                }
            }
            set
            {
                if (_ownerRun != null)
                {
                    _directRHld.Position = value;
                }
                else if (_ownerStyle != null)
                {
                    _directSHld.Position = value;
                }
                else if (_ownerParagraph != null)
                {
                    _markRHld.Position = value;
                }
                else if (_numberingLevel != null)
                {
                    _numRHld.Position = value;
                }
                else
                {
                    _position = value;
                }
            }
        }

        public Border Border
        {
            get => _border;
            internal set => _border = value;
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Clears all character formats.
        /// </summary>
        public void ClearFormatting()
        {
            if (_ownerRun != null)
            {
                _directRHld.ClearFormatting();
            }
            else if (_ownerStyle != null)
            {
                _directSHld.ClearFormatting();
            }
            else if(_ownerParagraph != null)
            {
                _markRHld.ClearFormatting();
            }
            else if(_numberingLevel != null)
            {
                _numRHld.ClearFormatting();
            }
        }
        #endregion
    }

}
