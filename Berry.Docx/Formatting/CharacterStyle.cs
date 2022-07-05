using System;
using System.Collections.Generic;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;

namespace Berry.Docx.Formatting
{
    public class CharacterStyle : Style
    {
        public CharacterStyle(Document doc):base(doc, StyleType.Character)
        {

        }
        internal CharacterStyle(Document doc, W.Style style) : base(doc, style)
        {

        }

        public CharacterFormat CharacterFormat => _cFormat;
        public new CharacterStyle BaseStyle
        {
            get => base.BaseStyle as CharacterStyle;
            set => base.BaseStyle = value;
        }

        public static CharacterStyle Default(Document doc)
        {
            return doc.Styles.Where(s => s.Type == StyleType.Character && s.IsDefault).FirstOrDefault() as CharacterStyle;
        }
    }
}
