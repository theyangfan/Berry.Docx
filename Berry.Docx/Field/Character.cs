using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

using Berry.Docx.Formatting;

namespace Berry.Docx.Field
{
    public class Character
    {
        private readonly TextRange _textRange;
        private readonly char _value;

        private string _fontName = string.Empty;
        private float _fontSize = 0;
        private bool _bold = false;
        private bool _italic = false;

        internal Character(TextRange tr, char value)
        {
            _textRange = tr;
            _value = value;

            var hint = tr.CharacterFormat.FontTypeHint;
            // complex script
            if (tr.CharacterFormat.UseComplexScript
                || tr.CharacterFormat.RightToLeft)
            {
                _fontName = tr.CharacterFormat.FontNameComplexScript;
                _fontSize = tr.CharacterFormat.FontSizeCs;
                _bold = tr.CharacterFormat.BoldCs;
                _italic = tr.CharacterFormat.ItalicCs;
            }
            else
            {
                // Basic Latin
                if (value <= 0x007F)
                {
                    _fontName = tr.CharacterFormat.FontNameAscii;
                }
                // Latin-1 Supplement
                else if(value >= 0x00A0 && value <= 0x00FF)
                {
                    if (hint == FontContentType.EastAsia
                        && (Regex.IsMatch(value.ToString(), @"[\xA1\xA4\xA7\xA8\xAA\xAD\xAF\xB0-\xB4\xB6-\xBA\xBC-\xBF\xD7\xF7]")
                        || Regex.IsMatch(value.ToString(), @"[\xE0\xE1\xE8-\xEA\xEC\xED\xF2\xF3\xF9-\xFA\xFC]")))
                        _fontName = tr.CharacterFormat.FontNameEastAsia;
                    else
                        _fontName = tr.CharacterFormat.FontNameHighAnsi;
                }
                // Latin Extended-A & Latin Extended-B, IPA Extensions
                else if (value >= 0x0100 && value <= 0x02AF)
                {
                    if (hint == FontContentType.EastAsia) _fontName = tr.CharacterFormat.FontNameEastAsia;
                    else _fontName = tr.CharacterFormat.FontNameHighAnsi;
                }
                // Spacing Modifier Letters, Combining Diacritical Marks, Greek & Cyrillic
                else if (value >= 0x02B0 && value <= 0x03CF
                    || value >= 0x0400 && value <= 0x04FF)
                {
                    if (hint == FontContentType.EastAsia) _fontName = tr.CharacterFormat.FontNameEastAsia;
                    else _fontName = tr.CharacterFormat.FontNameHighAnsi;
                }
                // Hebrew, Arabic, Syriac, Arabic Supplement, Thaana
                else if(value >= 0x0590 && value <= 0x07BF)
                {
                    _fontName = tr.CharacterFormat.FontNameAscii;
                }
                // Hangul Jamo
                else if(value >= 0x1100 && value <= 0x11FF)
                {
                    _fontName = tr.CharacterFormat.FontNameEastAsia;
                }
                // Latin Extended Additional
                else if (value >= 0x1E00 && value <= 0x1EFF)
                {
                    if (hint == FontContentType.EastAsia) _fontName = tr.CharacterFormat.FontNameEastAsia;
                    else _fontName = tr.CharacterFormat.FontNameHighAnsi;
                }
                // Greek Extended
                else if (value >= 0x1F00 && value <= 0x1FFF)
                {
                    _fontName = tr.CharacterFormat.FontNameHighAnsi;
                }
                // General Punctuation, Superscripts and Subscripts, Currency Symbols, Combining Diacritical 
                // Marks for Symbols, Letter-like Symbols, Number Forms, Arrows, Mathematical Operators,
                // Miscellaneous Technical, Control Pictures, Optical Character Recognition, Enclosed 
                // Alphanumerics, Box Drawing, Block Elements, Geometric Shapes, Miscellaneous Symbols,
                // Dingbats
                else if (value >= 0x2000 && value <= 0x27BF)
                {
                    if (hint == FontContentType.EastAsia) _fontName = tr.CharacterFormat.FontNameEastAsia;
                    else _fontName = tr.CharacterFormat.FontNameHighAnsi;
                }
                // CJK Radicals Supplement, Kangxi Radicals
                // Ideographic Description Characters, CJK Symbols and Punctuation, Hiragana, Katakana,
                // Bopomofo, Hangul Compatibility Jamo, Kanbun
                // Enclosed CJK Letters and Months, CJK Compatibility, CJK Unified Ideographs Extension A
                else if (value >= 0x2E80 && value <= 0x2FDF
                    || value >= 0x2FF0 && value <= 0x319F
                    || value >= 0x3200 && value <= 0x4DBF)
                {
                    _fontName = tr.CharacterFormat.FontNameEastAsia;
                }
                // CJK Unified Ideographs 
                else if (value >= 0x4E00 && value <= 0x9FAF)
                {
                    _fontName = tr.CharacterFormat.FontNameEastAsia;
                }
                // Yi Syllables, Yi Radicals, Hangul Syllables, High Surrogates, High Private Use Surrogates,
                // Low Surrogates
                else if (value >= 0xA000 && value <= 0xA4CF
                    || value >= 0xAC00 && value <= 0xD7AF
                    || value >= 0xD800 && value <= 0xDFFF)
                {
                    _fontName = tr.CharacterFormat.FontNameEastAsia;
                }
                // Private Use Area 
                else if (value >= 0xE000 && value <= 0xF8FF)
                {
                    if (hint == FontContentType.EastAsia) _fontName = tr.CharacterFormat.FontNameEastAsia;
                    else _fontName = tr.CharacterFormat.FontNameHighAnsi;
                }
                // CJK Compatibility Ideographs
                else if (value >= 0xF900 && value <= 0xFAFF)
                {
                    _fontName = tr.CharacterFormat.FontNameEastAsia;
                }
                // Alphabetic Presentation Forms
                else if (value >= 0xFB00 && value <= 0xFB4F)
                {
                    if (hint == FontContentType.EastAsia && value <= 0xFB1C)
                        _fontName = tr.CharacterFormat.FontNameEastAsia;
                    else if (value >= 0xFB1D)
                        _fontName = tr.CharacterFormat.FontNameAscii;
                    else
                        _fontName = tr.CharacterFormat.FontNameHighAnsi;
                }
                // Arabic Presentation Forms - A
                else if (value >= 0xFB50 && value <= 0xFDFF)
                {
                    _fontName = tr.CharacterFormat.FontNameAscii;
                }
                // CJK Compatibility Forms, Small Form Variants
                else if (value >= 0xFE30 && value <= 0xFE6F)
                {
                    _fontName = tr.CharacterFormat.FontNameEastAsia;
                }
                else if (value >= 0xFE70 && value <= 0xFEFE)
                {
                    _fontName = tr.CharacterFormat.FontNameAscii;
                }
                // Halfwidth and Fullwidth Forms
                else if (value >= 0xFF00 && value <= 0xFFFF)
                {
                    _fontName = tr.CharacterFormat.FontNameEastAsia;
                }

                _fontSize = tr.CharacterFormat.FontSize;
                _bold = tr.CharacterFormat.Bold;
                _italic = tr.CharacterFormat.Italic;
            }
        }

        public char Val => _value;

        public string FontName => _fontName;

        public float FontSize => _fontSize;

        public bool Bold => _bold;

        public bool Italic => _italic;

        public SubSuperScript SubSuperScript => _textRange.CharacterFormat.SubSuperScript;

        public UnderlineStyle UnderlineStyle => _textRange.CharacterFormat.UnderlineStyle;

        public ColorValue TextColor => _textRange.CharacterFormat.TextColor;

        public int CharacterScale => _textRange.CharacterFormat.CharacterScale;

        public float CharacterSpacing => _textRange.CharacterFormat.CharacterSpacing;

        public float Position => _textRange.CharacterFormat.Position;

        public bool IsHidden => _textRange.CharacterFormat.IsHidden;

        public bool SnapToGrid => _textRange.CharacterFormat.SnapToGrid;
    }
}
