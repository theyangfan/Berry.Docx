using P = DocumentFormat.OpenXml.Packaging;
using OOxml = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent an OpenXML RunProperties holder.
    /// </summary>
    internal class RunPropertiesHolder
    {
        #region Private Members
        private P.WordprocessingDocument _document;
        private OOxml.RunProperties _rPr = null;
        private OOxml.ParagraphMarkRunProperties _mark_rPr = null;
        private OOxml.StyleRunProperties _style_rPr = null;

        private OOxml.RunFonts _rFonts = null;
        private OOxml.FontSize _fontSize = null;
        private OOxml.FontSizeComplexScript _fontSizeCs = null;
        private OOxml.Bold _bold = null;
        private OOxml.Italic _italic = null;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the RunPropertiesHolder class using the supplied OpenXML RunProperties element.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="rPr"></param>
        public RunPropertiesHolder(P.WordprocessingDocument doc, OOxml.RunProperties rPr)
        {
            _document = doc;
            _rPr = rPr;
            _rFonts = rPr.RunFonts;
            _fontSize = rPr.FontSize;
            _fontSizeCs = rPr.FontSizeComplexScript;
            _bold = rPr.Bold;
            _italic = rPr.Italic;
        }

        /// <summary>
        /// Initializes a new instance of the RunPropertiesHolder class using the supplied OpenXML ParagraphMarkRunProperties element.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="rPr"></param>
        public RunPropertiesHolder(P.WordprocessingDocument doc, OOxml.ParagraphMarkRunProperties rPr)
        {
            _document = doc;
            _mark_rPr = rPr;
            _rFonts = rPr.GetFirstChild<OOxml.RunFonts>();
            _fontSize = rPr.GetFirstChild<OOxml.FontSize>();
            _fontSizeCs = rPr.GetFirstChild<OOxml.FontSizeComplexScript>();
            _bold = rPr.GetFirstChild<OOxml.Bold>();
            _italic = rPr.GetFirstChild<OOxml.Italic>();
        }

        /// <summary>
        /// Initializes a new instance of the RunPropertiesHolder class using the supplied OpenXML StyleRunProperties element.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="rPr"></param>
        public RunPropertiesHolder(P.WordprocessingDocument doc, OOxml.StyleRunProperties rPr)
        {
            _document = doc;
            _style_rPr = rPr;
            _rFonts = rPr.RunFonts;
            _fontSize = rPr.FontSize;
            _fontSizeCs = rPr.FontSizeComplexScript;
            _bold = rPr.Bold;
            _italic = rPr.Italic;
        }

        /// <summary>
        /// Initializes a new instance of the RunPropertiesHolder class using the supplied OpenXML RunPropertiesBaseStyle element.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="rPr"></param>
        public RunPropertiesHolder(P.WordprocessingDocument doc, OOxml.RunPropertiesBaseStyle rPr)
        {
            _document = doc;
            _rFonts = rPr.RunFonts;
            _fontSize = rPr.FontSize;
            _fontSizeCs = rPr.FontSizeComplexScript;
            _bold = rPr.Bold;
            _italic = rPr.Italic;
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
                if (_rFonts == null) return null;
                if(_rFonts.EastAsiaTheme != null)
                {
                    return _document.GetThemeFont(_rFonts.EastAsiaTheme);
                }
                return _rFonts.EastAsia;
            }
            set
            {
                if(_rFonts != null)
                {
                    _rFonts.EastAsia = value;
                }
                else
                {
                    _rFonts = new OOxml.RunFonts() { EastAsia = value };
                    if (_rPr != null) _rPr.RunFonts = _rFonts;
                    else if (_mark_rPr != null) _mark_rPr.AddChild(_rFonts);
                    else if (_style_rPr != null) _style_rPr.RunFonts = _rFonts;
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
                if (_rFonts == null) return null;
                if (_rFonts.AsciiTheme != null)
                {
                    return _document.GetThemeFont(_rFonts.AsciiTheme);
                }
                return _rFonts.Ascii;
            }
            set
            {
                if (_rFonts != null)
                {
                    _rFonts.Ascii = value;
                    _rFonts.HighAnsi = value;
                }
                else
                {
                    _rFonts = new OOxml.RunFonts() { Ascii = value, HighAnsi = value };
                    if (_rPr != null) _rPr.RunFonts = _rFonts;
                    else if (_mark_rPr != null) _mark_rPr.AddChild(_rFonts);
                    else if (_style_rPr != null) _style_rPr.RunFonts = _rFonts;
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
                if (_fontSize == null) return -1;
                return _fontSize.Val.Value.ToFloat() / 2;
            }
            set
            {
                if (_fontSize != null)
                {
                    _fontSize.Val = (value * 2).ToString();
                }
                else
                {
                    _fontSize = new OOxml.FontSize() { Val = (value*2).ToString() };
                    if (_rPr != null) _rPr.FontSize = _fontSize;
                    else if (_mark_rPr != null) _mark_rPr.AddChild(_fontSize);
                    else if (_style_rPr != null) _style_rPr.FontSize = _fontSize;
                }
            }
        }

        public float FontSizeCs
        {
            get
            {
                if (_fontSizeCs == null) return -1;
                return _fontSizeCs.Val.Value.ToFloat() / 2;
            }
            set
            {
                if (_fontSizeCs != null)
                {
                    _fontSizeCs.Val = (value * 2).ToString();
                }
                else
                {
                    _fontSizeCs = new OOxml.FontSizeComplexScript() { Val = (value * 2).ToString() };
                    if (_rPr != null) _rPr.FontSizeComplexScript = _fontSizeCs;
                    else if (_mark_rPr != null) _mark_rPr.AddChild(_fontSizeCs);
                    else if (_style_rPr != null) _style_rPr.FontSizeComplexScript = _fontSizeCs;
                }
            }
        }
        /// <summary>
        /// Gets or sets bold style.
        /// </summary>
        public ZBool Bold
        {
            get
            {
                if (_bold == null) return null;
                if (_bold.Val == null) return true;
                return _bold.Val.Value;
            }
            set
            {
                if(_bold != null)
                {
                    if (value) _bold.Val = null;
                    else _bold.Val = false;
                }
                else
                {
                    _bold = new OOxml.Bold();
                    if (value) _bold.Val = null;
                    else _bold.Val = false;
                    if (_rPr != null) _rPr.Bold = _bold;
                    else if (_mark_rPr != null) _mark_rPr.AddChild(_bold);
                    else if (_style_rPr != null) _style_rPr.Bold = _bold;
                }
            }
        }
        /// <summary>
        /// Gets or sets italic style.
        /// </summary>
        public ZBool Italic
        {
            get
            {
                if (_italic == null) return null;
                if (_italic.Val == null) return true;
                return _italic.Val.Value;
            }
            set
            {
                if (_italic != null)
                {
                    if (value) _italic.Val = null;
                    else _italic.Val = false;
                }
                else
                {
                    _italic = new OOxml.Italic();
                    if (value) _italic.Val = null;
                    else _italic.Val = false;
                    if (_rPr != null) _rPr.Italic = _italic;
                    else if (_mark_rPr != null) _mark_rPr.AddChild(_italic);
                    else if (_style_rPr != null) _style_rPr.Italic = _italic;
                }
            }
        }
        #endregion
    }
}
