using System;
using System.Collections.Generic;
using System.Text;

using Berry.Docx.Formatting;

namespace Berry.Docx.Field
{
    public class Character
    {
        private string _fontName = string.Empty;
        private float _fontSize = 0;
        private bool _bold = false;
        private bool _italic = false;

        public Character(TextRange ownerTextRange, char value)
        {
            FontNameType nameType = FontNameType.Ascii;
            // Basic Latin
            if(value <= 0x007F)
            {
                nameType = FontNameType.Ascii;
            }
        }

        public enum FontNameType
        {
            Ascii = 0,
            EastAsia = 1,
            HighAnsi = 2,
            ComplexScript = 3
        }

    }
}
