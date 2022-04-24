using System;
using System.Linq;
using System.Collections.Generic;
using OOxml = DocumentFormat.OpenXml.Wordprocessing;

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
        private OOxml.Run _ownerRun = null;
        private RunPropertiesHolder _curRHld = null;
        private CharacterFormat _rStyleFormat;
        private CharacterFormat _inheritFromParagraphFormat = null;
        #endregion

        #region Paragraph
        private OOxml.Paragraph _ownerParagraph = null;
        private CharacterFormat _pStyleFormat = null;
        #endregion

        #region ParagraphMark
        private OOxml.ParagraphMarkRunProperties _paragraphMark;
        private RunPropertiesHolder _curPMarkHld = null;
        #endregion

        #region Style
        private OOxml.Style _ownerStyle = null;
        private RunPropertiesHolder _curSHld = null;
        private CharacterFormat _inheritFromBaseStyleFormat = null;
        #endregion

        #region Formats
        private string _fontCN = "宋体";
        private string _fontEN = "Times New Roman";
        private float _fontSize = 10.5F;
        private float _fontSizeCs = 10.5F;
        private bool _bold = false;
        private bool _italic = false;
        private int _characterScale = 100;
        private float _characterSpacing = 0;
        private float _position = 0;
        #endregion

        #endregion

        #region Constructors
        internal CharacterFormat() { }

        /// <summary>
        /// Represent the character format of a TextRange.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ownerRun"></param>
        internal CharacterFormat(Document doc, OOxml.Run ownerRun)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            if (ownerRun.RunProperties == null)
                ownerRun.RunProperties = new OOxml.RunProperties();
            _curRHld = new RunPropertiesHolder(doc.Package, ownerRun.RunProperties);
            if (ownerRun.RunProperties.RunStyle != null)
                _rStyleFormat = new CharacterFormat(doc, ownerRun.GetStyle(doc));
            OOxml.Paragraph paragraph = ownerRun.Ancestors<OOxml.Paragraph>().FirstOrDefault();
            if (paragraph != null)
                _inheritFromParagraphFormat = new CharacterFormat(doc, paragraph);
        }

        /// <summary>
        /// Represent the character format of a Paragraph.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ownerParagraph"></param>
        internal CharacterFormat(Document doc, OOxml.Paragraph ownerParagraph)
        {
            _doc = doc;
            _ownerParagraph = ownerParagraph;
            _pStyleFormat = new CharacterFormat(doc, ownerParagraph.GetStyle(doc));
        }

        /// <summary>
        /// Represent the character format of a Style.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="ownerStyle"></param>
        internal CharacterFormat(Document doc, OOxml.Style ownerStyle)
        {
            _doc = doc;
            _ownerStyle = ownerStyle;
            if (ownerStyle.StyleRunProperties == null)
                ownerStyle.StyleRunProperties = new OOxml.StyleRunProperties();
            _curSHld = new RunPropertiesHolder(doc.Package, ownerStyle.StyleRunProperties);
            _inheritFromBaseStyleFormat = GetStyleCharacterFormatRecursively(ownerStyle);
        }

        internal CharacterFormat(Document doc, OOxml.ParagraphMarkRunProperties paragraphMark)
        {
            _doc = doc;
            _paragraphMark = paragraphMark;
            _curPMarkHld = new RunPropertiesHolder(doc.Package, paragraphMark);
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
                    InitRun();
                    if (_curRHld.FontNameEastAsia != null) return _curRHld.FontNameEastAsia;
                    if (_rStyleFormat?.FontNameEastAsia != null) return _rStyleFormat.FontNameEastAsia;
                    if (_inheritFromParagraphFormat?.FontNameEastAsia != null) return _inheritFromParagraphFormat.FontNameEastAsia;
                    return string.Empty;
                }
                else if(_ownerParagraph != null)
                {
                    return _pStyleFormat.FontNameEastAsia;
                }
                else if(_ownerStyle != null)
                {
                    return _curSHld.FontNameEastAsia ?? _inheritFromBaseStyleFormat.FontNameEastAsia;
                }
                else if(_paragraphMark != null)
                {
                    return _curPMarkHld.FontNameEastAsia ?? _fontCN;
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
                    _curRHld.FontNameEastAsia = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.FontNameEastAsia = value;
                }
                else if (_paragraphMark != null)
                {
                    _curPMarkHld.FontNameEastAsia = value;
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
                    InitRun();
                    if (_curRHld.FontNameAscii != null) return _curRHld.FontNameAscii;
                    if (_rStyleFormat?.FontNameAscii != null) return _rStyleFormat.FontNameAscii;
                    if (_inheritFromParagraphFormat?.FontNameAscii != null) return _inheritFromParagraphFormat.FontNameAscii;
                    return string.Empty;
                }
                else if (_ownerParagraph != null)
                {
                    return _pStyleFormat.FontNameAscii;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.FontNameAscii ?? _inheritFromBaseStyleFormat.FontNameAscii;
                }
                else if (_paragraphMark != null)
                {
                    return _curPMarkHld.FontNameAscii ?? _fontEN;
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
                    _curRHld.FontNameAscii = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.FontNameAscii = value;
                }
                else if (_paragraphMark != null)
                {
                    _curPMarkHld.FontNameAscii = value;
                }
                else
                {
                    _fontEN = value;
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
                    InitRun();
                    if (_curRHld.FontSize != null) return _curRHld.FontSize;
                    if (_rStyleFormat?.FontSize != null) return _rStyleFormat.FontSize;
                    if (_inheritFromParagraphFormat?.FontSize != null) return _inheritFromParagraphFormat.FontSize;
                    return 10.5F;
                }
                else if (_ownerParagraph != null)
                {
                    return _pStyleFormat.FontSize;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.FontSize ?? _inheritFromBaseStyleFormat.FontSize;
                }
                else if (_paragraphMark != null)
                {
                    return _curPMarkHld.FontSize ?? _fontSize;
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
                    _curRHld.FontSize = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.FontSize = value;
                }
                else if (_paragraphMark != null)
                {
                    _curPMarkHld.FontSize = value;
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
                if(_ownerRun != null)
                {
                    InitRun();
                    if (_curRHld.FontSizeCs != null) return _curRHld.FontSizeCs;
                    if (_rStyleFormat?.FontSizeCs != null) return _rStyleFormat.FontSizeCs;
                    if (_inheritFromParagraphFormat?.FontSizeCs != null) return _inheritFromParagraphFormat.FontSizeCs;
                    return 10.5F;
                }
                else if (_ownerParagraph != null)
                {
                    return _pStyleFormat.FontSizeCs;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.FontSizeCs ?? _inheritFromBaseStyleFormat.FontSizeCs;
                }
                else if (_paragraphMark != null)
                {
                    return _curPMarkHld.FontSizeCs ?? _fontSizeCs;
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
                    _curRHld.FontSizeCs = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.FontSizeCs = value;
                }
                else if (_paragraphMark != null)
                {
                    _curPMarkHld.FontSizeCs = value;
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
                if(_ownerRun != null)
                {
                    InitRun();
                    if (_curRHld.Bold != null) return _curRHld.Bold;
                    if(_inheritFromParagraphFormat != null)
                    {
                        if(_rStyleFormat != null)
                        {
                            return _rStyleFormat.Bold == _inheritFromParagraphFormat.Bold ? false : true;
                        }
                        return _inheritFromParagraphFormat.Bold;
                    }
                    return false;
                }
                else if (_ownerParagraph != null)
                {
                    return _pStyleFormat.Bold;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.Bold ?? _inheritFromBaseStyleFormat.Bold;
                }
                else if (_paragraphMark != null)
                {
                    return _curPMarkHld.Bold ?? false;
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
                    _curRHld.Bold = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.Bold = value;
                }
                else if (_paragraphMark != null)
                {
                    _curPMarkHld.Bold = value;
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
                if(_ownerRun != null)
                {
                    InitRun();
                    if (_curRHld.Italic != null) return _curRHld.Italic;
                    if (_inheritFromParagraphFormat != null)
                    {
                        if (_rStyleFormat != null)
                        {
                            return _rStyleFormat.Italic == _inheritFromParagraphFormat.Italic ? false : true;
                        }
                        return _inheritFromParagraphFormat.Italic;
                    }
                    return false;
                }
                if (_ownerParagraph != null)
                {
                    return _pStyleFormat.Italic;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.Italic ?? _inheritFromBaseStyleFormat.Italic;
                }
                else if (_paragraphMark != null)
                {
                    return _curPMarkHld.Italic ?? false;
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
                    _curRHld.Italic = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.Italic = value;
                }
                else if (_paragraphMark != null)
                {
                    _curPMarkHld.Italic = value;
                }
                else
                {
                    _italic = value;
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
                    InitRun();
                    if (_curRHld.CharacterScale != null) return _curRHld.CharacterScale;
                    if (_rStyleFormat?.CharacterScale != null) return _rStyleFormat.CharacterScale;
                    if (_inheritFromParagraphFormat?.CharacterScale != null) return _inheritFromParagraphFormat.CharacterScale;
                    return 100;
                }
                else if (_ownerParagraph != null)
                {
                    return _pStyleFormat.CharacterScale;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.CharacterScale ?? _inheritFromBaseStyleFormat.CharacterScale;
                }
                else if (_paragraphMark != null)
                {
                    return _curPMarkHld.CharacterScale ?? 100;
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
                    _curRHld.CharacterScale = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.CharacterScale = value;
                }
                else if (_paragraphMark != null)
                {
                    _curPMarkHld.CharacterScale = value;
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
                    InitRun();
                    if (_curRHld.CharacterSpacing != null) return _curRHld.CharacterSpacing;
                    if (_rStyleFormat?.CharacterSpacing != null) return _rStyleFormat.CharacterSpacing;
                    if (_inheritFromParagraphFormat?.CharacterSpacing != null) return _inheritFromParagraphFormat.CharacterSpacing;
                    return 0;
                }
                else if (_ownerParagraph != null)
                {
                    return _pStyleFormat.CharacterSpacing;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.CharacterSpacing ?? _inheritFromBaseStyleFormat.CharacterSpacing;
                }
                else if (_paragraphMark != null)
                {
                    return _curPMarkHld.CharacterSpacing ?? 0;
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
                    _curRHld.CharacterSpacing = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.CharacterSpacing = value;
                }
                else if (_paragraphMark != null)
                {
                    _curPMarkHld.CharacterSpacing = value;
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
                    InitRun();
                    if (_curRHld.Position != null) return _curRHld.Position;
                    if (_rStyleFormat?.Position != null) return _rStyleFormat.Position;
                    if (_inheritFromParagraphFormat?.Position != null) return _inheritFromParagraphFormat.Position;
                    return 0;
                }
                else if (_ownerParagraph != null)
                {
                    return _pStyleFormat.Position;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.Position ?? _inheritFromBaseStyleFormat.Position;
                }
                else if (_paragraphMark != null)
                {
                    return _curPMarkHld.Position ?? 0;
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
                    _curRHld.Position = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.Position = value;
                }
                else if (_paragraphMark != null)
                {
                    _curPMarkHld.Position = value;
                }
                else
                {
                    _position = value;
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
            if (_ownerRun != null)
            {
                _curRHld.clearFormatting();
            }
            else if (_ownerStyle != null)
            {
                _curSHld.clearFormatting();
            }
            else if (_paragraphMark != null)
            {
                _curPMarkHld.clearFormatting();
            }
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Returns the character format that specified in the style hierarchy of a style.
        /// </summary>
        /// <param name="style"> The style</param>
        /// <returns>The character format that specified in the style hierarchy.</returns> 
        private CharacterFormat GetStyleCharacterFormatRecursively(OOxml.Style style)
        {
            CharacterFormat format = new CharacterFormat();
            CharacterFormat baseFormat = new CharacterFormat();
            // Gets DOcDefaults
            OOxml.Styles styles = style.Parent as OOxml.Styles;
            if (styles.DocDefaults != null && styles.DocDefaults.RunPropertiesDefault != null
                && styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle != null)
            {
                RunPropertiesHolder rPr = new RunPropertiesHolder(_doc.Package, styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle);
                baseFormat.FontNameEastAsia = rPr.FontNameEastAsia;
                baseFormat.FontNameAscii = rPr.FontNameAscii;
                baseFormat.FontSize = rPr.FontSize ?? 10.5F;
                baseFormat.FontSizeCs = rPr.FontSizeCs ?? 10.5F;
                baseFormat.Bold = rPr.Bold ?? false;
                baseFormat.Italic = rPr.Italic ?? false;
                baseFormat.CharacterScale = rPr.CharacterScale ?? 100;
                baseFormat.CharacterSpacing = rPr.CharacterScale ?? 0;
                baseFormat.Position = rPr.Position ?? 0;
            }
            // Gets base style format
            OOxml.Style baseStyle = style.GetBaseStyle();
            if (baseStyle != null)
                baseFormat = GetStyleCharacterFormatRecursively(baseStyle);
            if (style.StyleRunProperties == null) style.StyleRunProperties = new OOxml.StyleRunProperties();
            RunPropertiesHolder curSHld = new RunPropertiesHolder(_doc.Package, style.StyleRunProperties);

            format.FontNameEastAsia = curSHld.FontNameEastAsia ?? baseFormat.FontNameEastAsia;
            format.FontNameAscii = curSHld.FontNameAscii ?? baseFormat.FontNameAscii;
            format.FontSize = curSHld.FontSize ?? baseFormat.FontSize;
            format.FontSizeCs = curSHld.FontSizeCs ?? baseFormat.FontSizeCs;
            format.Bold = curSHld.Bold ?? baseFormat.Bold;
            format.Italic = curSHld.Italic ?? baseFormat.Italic;
            format.CharacterScale = curSHld.CharacterScale ?? baseFormat.CharacterScale;
            format.CharacterSpacing = curSHld.CharacterSpacing ?? baseFormat.CharacterSpacing;
            format.Position = curSHld.Position ?? baseFormat.Position;
            return format;
        }
        private void InitRun()
        {
            if (_ownerRun.Parent != null && _inheritFromParagraphFormat == null)
                _inheritFromParagraphFormat = new CharacterFormat(_doc, _ownerRun.Parent as OOxml.Paragraph);
        }
        #endregion
    }

}
