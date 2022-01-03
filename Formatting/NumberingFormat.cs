using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OOxml = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    public class NumberingFormat
    {
        private OOxml.Level _lvl = null;

        public NumberingFormat(OOxml.Level lvl)
        {
            _lvl = lvl;
        }
        
        public int StartNumberingValue
        {
            get => _lvl.StartNumberingValue.Val;
        }

        public OOxml.NumberFormatValues NumberingType
        {
            get => _lvl.NumberingFormat.Val;
        }

        public string LevelText
        {
            get => _lvl.LevelText.Val;
        }

    }
}
