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
        public static OW.Body GetBody (this OP.WordprocessingDocument doc)
        {
            return doc?.MainDocumentPart?.Document?.Body;
        }

        public static OW.SectionProperties GetRootSectionProperties(this OP.WordprocessingDocument doc)
        {
            return doc.GetBody()?.LastChild as OW.SectionProperties;
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

        /// <summary>
        /// Get Theme Font.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="themeFont"></param>
        /// <returns></returns>
        public static string GetThemeFont(this OP.WordprocessingDocument doc, OW.ThemeFontValues themeFont)
        {
            OD.MajorFont majorFont = doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme?.MajorFont;
            OD.MinorFont minorFont = doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme?.MinorFont;
            OD.SupplementalFont font = null;
            switch (themeFont)
            {
                case OW.ThemeFontValues.MajorEastAsia:
                    font = majorFont.Elements<OD.SupplementalFont>().Where(f => f.Script.Value == "Hans").FirstOrDefault();
                    if(font != null)
                        return font.Typeface;
                    else
                        return majorFont.EastAsianFont.Typeface;
                case OW.ThemeFontValues.MajorAscii:
                    return majorFont.LatinFont.Typeface;
                case OW.ThemeFontValues.MajorHighAnsi:
                    return majorFont.LatinFont.Typeface;
                case OW.ThemeFontValues.MinorEastAsia:
                    font = minorFont.Elements<OD.SupplementalFont>().Where(f => f.Script.Value == "Hans").FirstOrDefault();
                    if (font != null)
                        return font.Typeface;
                    else
                        return minorFont.EastAsianFont.Typeface;
                case OW.ThemeFontValues.MinorAscii:
                    return minorFont.LatinFont.Typeface;
                case OW.ThemeFontValues.MinorHighAnsi:
                    return minorFont.LatinFont.Typeface;
                default:
                    return string.Empty;
            }
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
