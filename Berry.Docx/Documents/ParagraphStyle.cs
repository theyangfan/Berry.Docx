using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Formatting;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// TODO
    /// </summary>
    public class ParagraphStyle : Style
    {
        internal ParagraphStyle(Document doc, W.Style style):base(doc, style)
        {
        }

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
