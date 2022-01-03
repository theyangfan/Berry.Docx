using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Formatting;

namespace Berry.Docx.Documents
{
    public class ParagraphStyle : Style
    {
        public ParagraphStyle(W.Style style):base(style){}

        /// <summary>
        /// 段落格式
        /// </summary>
        public ParagraphFormat ParagraphFormat { get => _pFormat; }
        /// <summary>
        /// 字符格式
        /// </summary>
        public CharacterFormat CharacterFormat { get => _cFormat; }
    }
}
