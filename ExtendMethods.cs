using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using OOXML = DocumentFormat.OpenXml;
using OP = DocumentFormat.OpenXml.Packaging;
using OW = DocumentFormat.OpenXml.Wordprocessing;
using OD = DocumentFormat.OpenXml.Drawing;

namespace Berry.Docx
{
    static class ExtendMethods
    {
        public static OP.WordprocessingDocument Document(this OW.Paragraph p)
        {
            if (p == null) return null;
            DocumentFormat.OpenXml.Wordprocessing.Document document = p.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Document>().FirstOrDefault();
            if(document != null)
                return document.MainDocumentPart.OpenXmlPackage as OP.WordprocessingDocument;
            OW.Footnotes footnotes = p.Ancestors<OW.Footnotes>().FirstOrDefault();
            if (footnotes != null)
                return footnotes.FootnotesPart.OpenXmlPackage as OP.WordprocessingDocument;
            OW.Endnotes endnotes = p.Ancestors<OW.Endnotes>().FirstOrDefault();
            if (endnotes != null)
                return endnotes.EndnotesPart.OpenXmlPackage as OP.WordprocessingDocument;
            OW.Header header = p.Ancestors<OW.Header>().FirstOrDefault();
            if (header != null)
                return header.HeaderPart.OpenXmlPackage as OP.WordprocessingDocument;
            OW.Footer footer = p.Ancestors<OW.Footer>().FirstOrDefault();
            if (footer != null)
                return footer.FooterPart.OpenXmlPackage as OP.WordprocessingDocument;
            return null;
        }

        public static OP.WordprocessingDocument Document(this OW.ParagraphProperties pPr)
        {
            DocumentFormat.OpenXml.Wordprocessing.Document document = pPr.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Document>().First();
            return document.MainDocumentPart.OpenXmlPackage as OP.WordprocessingDocument;
        }

        public static OP.WordprocessingDocument Document(this OW.Style style)
        {
            DocumentFormat.OpenXml.Wordprocessing.Styles styles = style.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Styles>().First();
            return styles.StylesPart.OpenXmlPackage as OP.WordprocessingDocument;
        }


        public static OP.WordprocessingDocument Document(this OW.StyleParagraphProperties pPr)
        {
            DocumentFormat.OpenXml.Wordprocessing.Styles styles = pPr.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Styles>().First();
            return styles.StylesPart.OpenXmlPackage as OP.WordprocessingDocument;
        }

        public static OW.Style GetStyle(this OW.Paragraph p, Document doc)
        {
            OW.Styles styles = doc.Package.MainDocumentPart.StyleDefinitionsPart.Styles;
            if(p.ParagraphProperties != null && p.ParagraphProperties.ParagraphStyleId != null)
            {
                string styleId = p.ParagraphProperties.ParagraphStyleId.Val.ToString();
                return styles.Elements<OW.Style>().Where(s => s.StyleId == styleId).FirstOrDefault();
            }
            else
            {
                return styles.Elements<OW.Style>().Where(s => s.Type.Value == OW.StyleValues.Paragraph && s.Default != null && s.Default.Value == true).FirstOrDefault();
            }
        }

        public static OW.Style GetBaseStyle(this OW.Style style)
        {
            if(style.BasedOn != null)
            {
                OW.Styles styles = style.Parent as OW.Styles;
                string styleId = style.BasedOn.Val.ToString();
                return styles.Elements<OW.Style>().Where(s => s.StyleId == styleId).FirstOrDefault();
            }
            return null;
        }

        public static string RxReplace(this string input, string pattern, string newStr)
        {
            return Regex.Replace(input, pattern, newStr);
        }

        public static string GetMajorFont(this OP.WordprocessingDocument doc)
        {
            OD.Theme theme = doc.MainDocumentPart.ThemePart.Theme;
            if(theme != null && theme.ThemeElements != null 
                && theme.ThemeElements.FontScheme != null
                && theme.ThemeElements.FontScheme.MajorFont != null
                && theme.ThemeElements.FontScheme.MajorFont.LatinFont != null
                && theme.ThemeElements.FontScheme.MajorFont.LatinFont.Typeface != null)
            {
                return theme.ThemeElements.FontScheme.MajorFont.LatinFont.Typeface;
            }
            return "宋体";
        }

        public static string GetMinorFont(this OP.WordprocessingDocument doc)
        {
            OD.Theme theme = doc.MainDocumentPart.ThemePart.Theme;
            if (theme != null && theme.ThemeElements != null
                && theme.ThemeElements.FontScheme != null
                && theme.ThemeElements.FontScheme.MinorFont != null
                && theme.ThemeElements.FontScheme.MinorFont.LatinFont != null
                && theme.ThemeElements.FontScheme.MinorFont.LatinFont.Typeface != null)
            {
                return theme.ThemeElements.FontScheme.MinorFont.LatinFont.Typeface;
            }
            return "宋体";
        }
        /// <summary>
        /// 转换为单精度浮点数
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static float ToFloat(this string str)
        {
            float val = 0;
            float.TryParse(str, out val);
            return val;
        }
        /// <summary>
        /// 转换为整型
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static int ToInt(this string str)
        {
            int val = 0;
            int.TryParse(str, out val);
            return val;
        }

        #region Enum Converter
        public static JustificationType Convert(this OW.JustificationValues type)
        {
            switch (type)
            {
                case OW.JustificationValues.Left:
                    return JustificationType.Left;
                case OW.JustificationValues.Center:
                    return JustificationType.Center;
                case OW.JustificationValues.Right:
                    return JustificationType.Right;
                case OW.JustificationValues.Both:
                    return JustificationType.Both;
                case OW.JustificationValues.Distribute:
                    return JustificationType.Distribute;
                default:
                    return JustificationType.None;
            }
        }
        public static OW.JustificationValues Convert(this JustificationType type)
        {
            switch (type)
            {
                case JustificationType.Left:
                    return OW.JustificationValues.Left;
                case JustificationType.Center:
                    return OW.JustificationValues.Center;
                case JustificationType.Right:
                    return OW.JustificationValues.Right;
                case JustificationType.Both:
                    return OW.JustificationValues.Both;
                case JustificationType.Distribute:
                    return OW.JustificationValues.Distribute;
                default:
                    return OW.JustificationValues.Both;
            }
        }
        #endregion
    }
}
