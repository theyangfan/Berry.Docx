using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

using Berry.Docx.Formatting;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represent the character in the paragraph.
    /// </summary>
    public class Character
    {
        #region Private Members
        private readonly char _value;
        private readonly CharacterFormat _format;

        private string _fontName = string.Empty;
        private float _fontSize = 0;
        private bool _bold = false;
        private bool _italic = false;
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new <see cref="Character"/> instance.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="format"></param>
        public Character(char value, CharacterFormat format)
        {
            _value = value;
            _format = format;

            var hint = format.FontTypeHint;
            // complex script
            if (format.UseComplexScript
                || format.RightToLeft)
            {
                _fontName = format.FontNameComplexScript;
                _fontSize = format.FontSizeCs;
                _bold = format.BoldCs;
                _italic = format.ItalicCs;
            }
            else
            {
                // Basic Latin
                if (value <= 0x007F)
                {
                    _fontName = format.FontNameAscii;
                }
                // Latin-1 Supplement
                else if(value >= 0x00A0 && value <= 0x00FF)
                {
                    if (hint == FontContentType.EastAsia
                        && (Regex.IsMatch(value.ToString(), @"[\xA1\xA4\xA7\xA8\xAA\xAD\xAF\xB0-\xB4\xB6-\xBA\xBC-\xBF\xD7\xF7]")
                        || Regex.IsMatch(value.ToString(), @"[\xE0\xE1\xE8-\xEA\xEC\xED\xF2\xF3\xF9-\xFA\xFC]")))
                        _fontName = format.FontNameEastAsia;
                    else
                        _fontName = format.FontNameHighAnsi;
                }
                // Latin Extended-A & Latin Extended-B, IPA Extensions
                else if (value >= 0x0100 && value <= 0x02AF)
                {
                    if (hint == FontContentType.EastAsia) _fontName = format.FontNameEastAsia;
                    else _fontName = format.FontNameHighAnsi;
                }
                // Spacing Modifier Letters, Combining Diacritical Marks, Greek & Cyrillic
                else if (value >= 0x02B0 && value <= 0x03CF
                    || value >= 0x0400 && value <= 0x04FF)
                {
                    if (hint == FontContentType.EastAsia) _fontName = format.FontNameEastAsia;
                    else _fontName = format.FontNameHighAnsi;
                }
                // Hebrew, Arabic, Syriac, Arabic Supplement, Thaana
                else if(value >= 0x0590 && value <= 0x07BF)
                {
                    _fontName = format.FontNameAscii;
                }
                // Hangul Jamo
                else if(value >= 0x1100 && value <= 0x11FF)
                {
                    _fontName = format.FontNameEastAsia;
                }
                // Latin Extended Additional
                else if (value >= 0x1E00 && value <= 0x1EFF)
                {
                    if (hint == FontContentType.EastAsia) _fontName = format.FontNameEastAsia;
                    else _fontName = format.FontNameHighAnsi;
                }
                // Greek Extended
                else if (value >= 0x1F00 && value <= 0x1FFF)
                {
                    _fontName = format.FontNameHighAnsi;
                }
                // General Punctuation, Superscripts and Subscripts, Currency Symbols, Combining Diacritical 
                // Marks for Symbols, Letter-like Symbols, Number Forms, Arrows, Mathematical Operators,
                // Miscellaneous Technical, Control Pictures, Optical Character Recognition, Enclosed 
                // Alphanumerics, Box Drawing, Block Elements, Geometric Shapes, Miscellaneous Symbols,
                // Dingbats
                else if (value >= 0x2000 && value <= 0x27BF)
                {
                    if (hint == FontContentType.EastAsia) _fontName = format.FontNameEastAsia;
                    else _fontName = format.FontNameHighAnsi;
                }
                // CJK Radicals Supplement, Kangxi Radicals
                // Ideographic Description Characters, CJK Symbols and Punctuation, Hiragana, Katakana,
                // Bopomofo, Hangul Compatibility Jamo, Kanbun
                // Enclosed CJK Letters and Months, CJK Compatibility, CJK Unified Ideographs Extension A
                else if (value >= 0x2E80 && value <= 0x2FDF
                    || value >= 0x2FF0 && value <= 0x319F
                    || value >= 0x3200 && value <= 0x4DBF)
                {
                    _fontName = format.FontNameEastAsia;
                }
                // CJK Unified Ideographs 
                else if (value >= 0x4E00 && value <= 0x9FAF)
                {
                    _fontName = format.FontNameEastAsia;
                }
                // Yi Syllables, Yi Radicals, Hangul Syllables, High Surrogates, High Private Use Surrogates,
                // Low Surrogates
                else if (value >= 0xA000 && value <= 0xA4CF
                    || value >= 0xAC00 && value <= 0xD7AF
                    || value >= 0xD800 && value <= 0xDFFF)
                {
                    _fontName = format.FontNameEastAsia;
                }
                // Private Use Area 
                else if (value >= 0xE000 && value <= 0xF8FF)
                {
                    if (hint == FontContentType.EastAsia) _fontName = format.FontNameEastAsia;
                    else _fontName = format.FontNameHighAnsi;
                }
                // CJK Compatibility Ideographs
                else if (value >= 0xF900 && value <= 0xFAFF)
                {
                    _fontName = format.FontNameEastAsia;
                }
                // Alphabetic Presentation Forms
                else if (value >= 0xFB00 && value <= 0xFB4F)
                {
                    if (hint == FontContentType.EastAsia && value <= 0xFB1C)
                        _fontName = format.FontNameEastAsia;
                    else if (value >= 0xFB1D)
                        _fontName = format.FontNameAscii;
                    else
                        _fontName = format.FontNameHighAnsi;
                }
                // Arabic Presentation Forms - A
                else if (value >= 0xFB50 && value <= 0xFDFF)
                {
                    _fontName = format.FontNameAscii;
                }
                // CJK Compatibility Forms, Small Form Variants
                else if (value >= 0xFE30 && value <= 0xFE6F)
                {
                    _fontName = format.FontNameEastAsia;
                }
                else if (value >= 0xFE70 && value <= 0xFEFE)
                {
                    _fontName = format.FontNameAscii;
                }
                // Halfwidth and Fullwidth Forms
                else if (value >= 0xFF00 && value <= 0xFFFF)
                {
                    _fontName = format.FontNameEastAsia;
                }

                _fontSize = format.FontSize;
                _bold = format.Bold;
                _italic = format.Italic;
            }
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The char value.
        /// </summary>
        public char Val => _value;

        /// <summary>
        /// Gets the font family name.
        /// </summary>
        public string FontName => _fontName;

        /// <summary>
        /// Gets the font size.
        /// </summary>
        public float FontSize => _fontSize;

        /// <summary>
        /// Bold font weight.
        /// </summary>
        public bool Bold => _bold;

        /// <summary>
        /// Italic font style.
        /// </summary>
        public bool Italic => _italic;

        /// <summary>
        /// Gets the vertical positioning of the character.
        /// </summary>
        public SubSuperScript SubSuperScript => _format.SubSuperScript;

        /// <summary>
        /// Gets the underline style.
        /// </summary>
        public UnderlineStyle UnderlineStyle => _format.UnderlineStyle;

        /// <summary>
        /// Gets the text color.
        /// </summary>
        public ColorValue TextColor => _format.TextColor;

        /// <summary>
        /// Gets the percent value of the normal character width that each 
        /// character shall be scaled.
        /// </summary>
        public int CharacterScale => _format.CharacterScale;

        /// <summary>
        /// Gets the amount (in points) of character pitch which shall be added 
        /// or removed after each character.
        /// </summary>
        public float CharacterSpacing => _format.CharacterSpacing;

        /// <summary>
        /// Gets the amount (in points) by which text shall be raised 
        /// or lowered in relation to the default baseline location.
        /// </summary>
        public float Position => _format.Position;

        /// <summary>
        /// Gets a value indicating whether the text is hidden.
        /// </summary>
        public bool IsHidden => _format.IsHidden;

        /// <summary>
        /// Gets a value indicating whether snap to grid when document grid is defined.
        /// </summary>
        public bool SnapToGrid => _format.SnapToGrid;
        #endregion
    }
}
