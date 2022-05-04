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
    /// 
    /// </summary>
    public class ParagraphStyle : Style
    {
        private readonly Document _doc;
        public ParagraphStyle(Document doc) : base(doc, StyleType.Paragraph)
        {
            _doc = doc;
        }
        internal ParagraphStyle(Document doc, W.Style style):base(doc, style)
        {
            _doc = doc;
        }

        /// <summary>
        /// 段落格式
        /// </summary>
        public ParagraphFormat ParagraphFormat { get => _pFormat; }
        /// <summary>
        /// 字符格式
        /// </summary>
        public CharacterFormat CharacterFormat { get => _cFormat; }

        public new ParagraphStyle BaseStyle
        {
            get => base.BaseStyle as ParagraphStyle;
            set => base.BaseStyle = value;
        }

        public static ParagraphStyle Default(Document doc)
        {
            return doc.Styles.Where(s => s.Type == StyleType.Paragraph && s.IsDefault).FirstOrDefault() as ParagraphStyle;
        }

        public CharacterStyle GetLinkedStyle()
        {
            return LinkedStyle as CharacterStyle;
        }

        public CharacterStyle CreateLinkedStyle()
        {
            if(LinkedStyle != null)
            {
                return LinkedStyle as CharacterStyle;
            }
            CharacterStyle linked = new CharacterStyle(_doc);
            linked.Name = this.Name + " 字符";
            linked.BaseStyle = CharacterStyle.Default(_doc);
            linked.LinkedStyle = this;
            linked.IsCustom = true;
            this.LinkedStyle = linked;
            _doc.Styles.Add(linked);
            return linked;
        }

    }
}
