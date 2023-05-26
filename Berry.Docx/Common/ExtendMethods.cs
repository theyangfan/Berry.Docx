using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using P = DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

using Berry.Docx.Documents;

namespace Berry.Docx
{
    internal static class ExtendMethods
    {
        /// <summary>
        /// Converts string to float.
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
        /// Converts string to int.
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static int ToInt(this string str)
        {
            int val = 0;
            int.TryParse(str, out val);
            return val;
        }

        public static float Round(this float val, int decimals)
        {
            return (float)Math.Round(val, decimals);
        }

        /// <summary>
        /// In current string, replaces all strings that match a specified regular
        /// expression with a specified replacement string.
        /// </summary>
        /// <param name="input">The current string.</param>
        /// <param name="pattern">The regular expression pattern to match.</param>
        /// <param name="newStr">The replacement string.</param>
        /// <returns>
        /// A new string that is identical to the input string, except that the replacement
        /// string takes the place of each matched string. If pattern is not matched in the
        /// current instance, the method returns the current instance unchanged.
        /// </returns>
        public static string RxReplace(this string input, string pattern, string newStr)
        {
            return Regex.Replace(input, pattern, newStr);
        }

#if NET35
        public static bool HasFlag(this TextReadingMode e, TextReadingMode flag)
        {
            byte a = (byte)e;
            byte b = (byte)flag;
            return (a & b) == b;
        }

        public static IEnumerable<DocumentItem> Convert<T>(this IEnumerable<T> items) where T : DocumentItem
        {
            foreach(var item in items)
            {
                yield return item;
            }
        }

        public static IEnumerable<DocumentObject> Convert(this IEnumerable<DocumentItem> items)
        {
            foreach (var item in items)
            {
                yield return item;
            }
        }
#endif

        #region OpenXMl Extend Methods
        /// <summary>
        /// Returns document body element.
        /// </summary>
        /// <param name="doc">The OpenXML WordprocessingDocument</param>
        /// <returns>The body element.</returns>
        internal static W.Body GetBody (this P.WordprocessingDocument doc)
        {
            return doc?.MainDocumentPart?.Document?.Body;
        }

        /// <summary>
        /// Returns the root(last) child SectionProperties element of Document. 
        /// </summary>
        /// <param name="doc">The OpenXML WordprocessingDocument</param>
        /// <returns>The SectionProperties</returns>
        internal static W.SectionProperties GetRootSectionProperties(this P.WordprocessingDocument doc)
        {
            return doc.GetBody()?.LastChild as W.SectionProperties;
        }

        /// <summary>
        /// Returns the OpenXML style that referenced by the paragraph.
        /// </summary>
        /// <param name="p">The OpenXMl paragraph element.</param>
        /// <param name="doc">The document</param>
        /// <returns>The OpenXML style</returns>
        internal static W.Style GetStyle(this W.Paragraph p, Document doc)
        {
            W.Styles styles = doc.Package.MainDocumentPart.StyleDefinitionsPart.Styles;
            if(p?.ParagraphProperties?.ParagraphStyleId != null)
            {
                string styleId = p.ParagraphProperties.ParagraphStyleId.Val.ToString();
                return styles.Elements<W.Style>().Where(s => s.StyleId == styleId).FirstOrDefault();
            }
            else
            {
                return styles.Elements<W.Style>().Where(s => s.Type.Value == W.StyleValues.Paragraph &&  s.Default?.Value == true).FirstOrDefault();
            }
        }

        internal static W.Style GetStyle(this W.Run run, Document doc)
        {
            W.Styles styles = doc.Package.MainDocumentPart.StyleDefinitionsPart.Styles;
            if (run?.RunProperties?.RunStyle != null)
            {
                string styleId = run.RunProperties.RunStyle.Val.ToString();
                return styles.Elements<W.Style>().Where(s => s.StyleId == styleId).FirstOrDefault();
            }
            return null;
        }

        /// <summary>
        /// Returns the OpenXML style that the current style based on.
        /// </summary>
        /// <param name="style">The OpenXMl style.</param>
        /// <returns>The based-on OpenXMl style.</returns>
        internal static W.Style GetBaseStyle(this W.Style style, Document doc)
        {
            if(style.BasedOn != null)
            {
                W.Styles styles = doc.Package.MainDocumentPart.StyleDefinitionsPart.Styles;
                string styleId = style.BasedOn.Val.ToString();
                return styles.Elements<W.Style>().Where(s => s.StyleId == styleId).FirstOrDefault();
            }
            return null;
        }
        
        /// <summary>
        /// Get Theme Font.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="themeFont"></param>
        /// <returns></returns>
        public static string GetThemeFont(this P.WordprocessingDocument doc, W.ThemeFontValues themeFont)
        {
            Dictionary<W.ThemeFontValues, string> themeFonts = new Dictionary<W.ThemeFontValues, string>();
            A.FontScheme fonts = doc.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme;
            if (fonts != null)
            {
                var majorFont = fonts.MajorFont;
                var minorFont = fonts.MinorFont;
                if (majorFont != null)
                {
                    var latin = majorFont.LatinFont;
                    var eastAsian = majorFont.EastAsianFont;
                    var cs = majorFont.ComplexScriptFont;
                    var hans = majorFont.Elements<A.SupplementalFont>().Where(f => f.Script == "Hans").FirstOrDefault();
                    var arab = majorFont.Elements<A.SupplementalFont>().Where(f => f.Script == "Arab").FirstOrDefault();
                    if (!string.IsNullOrEmpty(latin?.Typeface)) 
                    {
                        themeFonts.Add(W.ThemeFontValues.MajorAscii, latin.Typeface);
                        themeFonts.Add(W.ThemeFontValues.MajorHighAnsi, latin.Typeface);
                    }
                    if (!string.IsNullOrEmpty(eastAsian?.Typeface)) themeFonts.Add(W.ThemeFontValues.MajorEastAsia, eastAsian.Typeface);
                    else if (!string.IsNullOrEmpty(hans?.Typeface)) themeFonts.Add(W.ThemeFontValues.MajorEastAsia, hans.Typeface);
                    if (!string.IsNullOrEmpty(cs?.Typeface)) themeFonts.Add(W.ThemeFontValues.MajorBidi, cs.Typeface);
                    else if (!string.IsNullOrEmpty(arab?.Typeface)) themeFonts.Add(W.ThemeFontValues.MajorBidi, arab.Typeface);
                }
                if (minorFont != null)
                {
                    var latin = minorFont.LatinFont;
                    var eastAsian = minorFont.EastAsianFont;
                    var cs = minorFont.ComplexScriptFont;
                    var hans = minorFont.Elements<A.SupplementalFont>().Where(f => f.Script == "Hans").FirstOrDefault();
                    var arab = minorFont.Elements<A.SupplementalFont>().Where(f => f.Script == "Arab").FirstOrDefault();
                    if (!string.IsNullOrEmpty(latin?.Typeface))
                    {
                        themeFonts.Add(W.ThemeFontValues.MinorAscii, latin.Typeface);
                        themeFonts.Add(W.ThemeFontValues.MinorHighAnsi, latin.Typeface);
                    }
                    if (!string.IsNullOrEmpty(eastAsian?.Typeface)) themeFonts.Add(W.ThemeFontValues.MinorEastAsia, eastAsian.Typeface);
                    else if (!string.IsNullOrEmpty(hans?.Typeface)) themeFonts.Add(W.ThemeFontValues.MinorEastAsia, hans.Typeface);
                    if (!string.IsNullOrEmpty(cs?.Typeface)) themeFonts.Add(W.ThemeFontValues.MinorBidi, cs.Typeface);
                    else if (!string.IsNullOrEmpty(arab?.Typeface)) themeFonts.Add(W.ThemeFontValues.MinorBidi, arab.Typeface);
                }
            }
            return themeFonts.ContainsKey(themeFont) ? themeFonts[themeFont] : string.Empty;
        }
#endregion
    }
}
