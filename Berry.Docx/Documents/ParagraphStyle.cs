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

        public ParagraphStyle(Document doc, string styleName) : this(doc, styleName, BuiltInStyle.Normal)
        {
        }

        public ParagraphStyle(Document doc, string styleName, BuiltInStyle basedStyle) : base(doc, StyleType.Paragraph)
        {
            _doc = doc;
            this.Name = styleName;
            this.IsCustom = true;
            this.AddToGallery = true;
            if(basedStyle != BuiltInStyle.None)
            {
                this.BaseStyle = CreateBuiltInStyle(basedStyle, doc);
            }
        }

        internal ParagraphStyle(Document doc, W.Style style):base(doc, style)
        {
            _doc = doc;
        }

        /// <summary>
        /// 段落格式
        /// </summary>
        public ParagraphFormat ParagraphFormat => _pFormat;
        /// <summary>
        /// 字符格式
        /// </summary>
        public CharacterFormat CharacterFormat => _cFormat;

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
            return base.LinkedStyle as CharacterStyle;
        }

        public CharacterStyle CreateLinkedStyle()
        {
            if(base.LinkedStyle != null)
            {
                return base.LinkedStyle as CharacterStyle;
            }
            CharacterStyle linked = new CharacterStyle(_doc);
            linked.Name = this.Name.Replace("heading", "标题") + " 字符";
            linked.BaseStyle = CharacterStyle.Default(_doc);
            linked.LinkedStyle = this;
            linked.IsCustom = true;
            this.LinkedStyle = linked;
            // copy character format
            linked.CharacterFormat.FontNameEastAsia = this.CharacterFormat.FontNameEastAsia;
            linked.CharacterFormat.FontNameAscii = this.CharacterFormat.FontNameAscii;
            linked.CharacterFormat.FontSize = this.CharacterFormat.FontSize;
            linked.CharacterFormat.FontSizeCs = this.CharacterFormat.FontSizeCs;
            linked.CharacterFormat.Bold = this.CharacterFormat.Bold;
            linked.CharacterFormat.Italic = this.CharacterFormat.Italic;
            linked.CharacterFormat.CharacterScale = this.CharacterFormat.CharacterScale;
            linked.CharacterFormat.CharacterSpacing = this.CharacterFormat.CharacterSpacing;
            linked.CharacterFormat.Position = this.CharacterFormat.Position;
            // add to style list
            _doc.Styles.Add(linked);
            return linked;
        }

        public static ParagraphStyle CreateBuiltInStyle(BuiltInStyle bstyle, Document doc)
        {
            string styleName = string.Empty;
            switch (bstyle)
            {
                case BuiltInStyle.Normal:
                    styleName = "Normal";
                    break;
                case BuiltInStyle.Heading1:
                    styleName = "heading 1";
                    break;
                case BuiltInStyle.Heading2:
                    styleName = "heading 2";
                    break;
                case BuiltInStyle.Heading3:
                    styleName = "heading 3";
                    break;
                case BuiltInStyle.Heading4:
                    styleName = "heading 4";
                    break;
                case BuiltInStyle.Heading5:
                    styleName = "heading 5";
                    break;
                case BuiltInStyle.Heading6:
                    styleName = "heading 6";
                    break;
                case BuiltInStyle.Heading7:
                    styleName = "heading 7";
                    break;
                case BuiltInStyle.Heading8:
                    styleName = "heading 8";
                    break;
                case BuiltInStyle.Heading9:
                    styleName = "heading 9";
                    break;
                default:
                    break;
            }
            if (string.IsNullOrEmpty(styleName)) return null;

            if (doc.Styles.FindByName(styleName, StyleType.Paragraph) != null)
            {
                return doc.Styles.FindByName(styleName, StyleType.Paragraph) as ParagraphStyle;
            }
            else
            {
                W.Style s = BuiltInStyleGenerator.Generate(doc, bstyle);
                if(s == null) return null;
                ParagraphStyle style = new ParagraphStyle(doc, s);
                doc.Styles.Add(style);
                return style;
            }
        }

        


    }
}
