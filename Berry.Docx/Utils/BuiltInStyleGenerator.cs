using System;
using System.Collections.Generic;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx
{
    internal class BuiltInStyleGenerator
    {
        public static Style Generate(Document doc, BuiltInStyle style)
        {
            switch (style)
            {
                case BuiltInStyle.Normal:
                    return GenerateNormal(doc);
                case BuiltInStyle.Heading1:
                    return GenerateHeading1(doc);
                case BuiltInStyle.Heading2:
                    return GenerateHeading2(doc);
                case BuiltInStyle.Heading3:
                    return GenerateHeading3(doc);
                case BuiltInStyle.Heading4:
                    return GenerateHeading4(doc);
                case BuiltInStyle.Heading5:
                    return GenerateHeading5(doc);
                case BuiltInStyle.Heading6:
                    return GenerateHeading6(doc);
                case BuiltInStyle.Heading7:
                    return GenerateHeading7(doc);
                case BuiltInStyle.Heading8:
                    return GenerateHeading8(doc);
                case BuiltInStyle.Heading9:
                    return GenerateHeading9(doc);
                case BuiltInStyle.TOC1:
                    return GenerateTOC1(doc);
                case BuiltInStyle.TOC2:
                    return GenerateTOC2(doc);
                case BuiltInStyle.TOC3:
                    return GenerateTOC3(doc);
                default:
                    return null;
            }
        }

        private static Style GenerateNormal(Document doc)
        {
            string id = IDGenerator.GenerateStyleID(doc);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = id, Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            WidowControl widowControl1 = new WidowControl() { Val = false };
            Justification justification1 = new Justification() { Val = JustificationValues.Both };

            styleParagraphProperties1.Append(widowControl1);
            styleParagraphProperties1.Append(justification1);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            return style1;
        }

        private static Style GenerateHeading1(Document doc)
        {
            string id = IDGenerator.GenerateStyleID(doc);
            string baseId = Berry.Docx.Documents.ParagraphStyle.CreateBuiltInStyle(BuiltInStyle.Normal, doc).StyleId;

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = id };
            StyleName styleName1 = new StyleName() { Val = "heading 1" };
            BasedOn basedOn1 = new BasedOn() { Val = baseId };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = baseId };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "340", After = "330", Line = "578", LineRule = LineSpacingRuleValues.Auto };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 0 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            Kern kern1 = new Kern() { Val = (UInt32Value)44U };
            FontSize fontSize1 = new FontSize() { Val = "44" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "44" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(kern1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        private static Style GenerateHeading2(Document doc)
        {
            string id = IDGenerator.GenerateStyleID(doc);
            string baseId = Berry.Docx.Documents.ParagraphStyle.CreateBuiltInStyle(BuiltInStyle.Normal, doc).StyleId;

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = id };
            StyleName styleName1 = new StyleName() { Val = "heading 2" };
            BasedOn basedOn1 = new BasedOn() { Val = baseId };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = baseId };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "260", After = "260", Line = "416", LineRule = LineSpacingRuleValues.Auto };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 1 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        private static Style GenerateHeading3(Document doc)
        {
            string id = IDGenerator.GenerateStyleID(doc);
            string baseId = Berry.Docx.Documents.ParagraphStyle.CreateBuiltInStyle(BuiltInStyle.Normal, doc).StyleId;
            
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = id };
            StyleName styleName1 = new StyleName() { Val = "heading 3" };
            BasedOn basedOn1 = new BasedOn() { Val = baseId };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = baseId };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "260", After = "260", Line = "416", LineRule = LineSpacingRuleValues.Auto };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 2 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        private static Style GenerateHeading4(Document doc)
        {
            string id = IDGenerator.GenerateStyleID(doc);
            string baseId = Berry.Docx.Documents.ParagraphStyle.CreateBuiltInStyle(BuiltInStyle.Normal, doc).StyleId;

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = id };
            StyleName styleName1 = new StyleName() { Val = "heading 4" };
            BasedOn basedOn1 = new BasedOn() { Val = baseId };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = baseId };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "280", After = "290", Line = "376", LineRule = LineSpacingRuleValues.Auto };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 3 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        private static Style GenerateHeading5(Document doc)
        {
            string id = IDGenerator.GenerateStyleID(doc);
            string baseId = Berry.Docx.Documents.ParagraphStyle.CreateBuiltInStyle(BuiltInStyle.Normal, doc).StyleId;

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = id };
            StyleName styleName1 = new StyleName() { Val = "heading 5" };
            BasedOn basedOn1 = new BasedOn() { Val = baseId };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = baseId };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "280", After = "290", Line = "376", LineRule = LineSpacingRuleValues.Auto };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 4 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        private static Style GenerateHeading6(Document doc)
        {
            string id = IDGenerator.GenerateStyleID(doc);
            string baseId = Berry.Docx.Documents.ParagraphStyle.CreateBuiltInStyle(BuiltInStyle.Normal, doc).StyleId;

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = id };
            StyleName styleName1 = new StyleName() { Val = "heading 6" };
            BasedOn basedOn1 = new BasedOn() { Val = baseId };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = baseId };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240", After = "64", Line = "320", LineRule = LineSpacingRuleValues.Auto };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 5 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        private static Style GenerateHeading7(Document doc)
        {
            string id = IDGenerator.GenerateStyleID(doc);
            string baseId = Berry.Docx.Documents.ParagraphStyle.CreateBuiltInStyle(BuiltInStyle.Normal, doc).StyleId;

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = id };
            StyleName styleName1 = new StyleName() { Val = "heading 7" };
            BasedOn basedOn1 = new BasedOn() { Val = baseId };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = baseId };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240", After = "64", Line = "320", LineRule = LineSpacingRuleValues.Auto };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 6 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        private static Style GenerateHeading8(Document doc)
        {
            string id = IDGenerator.GenerateStyleID(doc);
            string baseId = Berry.Docx.Documents.ParagraphStyle.CreateBuiltInStyle(BuiltInStyle.Normal, doc).StyleId;

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = id };
            StyleName styleName1 = new StyleName() { Val = "heading 8" };
            BasedOn basedOn1 = new BasedOn() { Val = baseId };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = baseId };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240", After = "64", Line = "320", LineRule = LineSpacingRuleValues.Auto };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 7 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        private static Style GenerateHeading9(Document doc)
        {
            string id = IDGenerator.GenerateStyleID(doc);
            string baseId = Berry.Docx.Documents.ParagraphStyle.CreateBuiltInStyle(BuiltInStyle.Normal, doc).StyleId;

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = id };
            StyleName styleName1 = new StyleName() { Val = "heading 9" };
            BasedOn basedOn1 = new BasedOn() { Val = baseId };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = baseId };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240", After = "64", Line = "320", LineRule = LineSpacingRuleValues.Auto };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 8 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };

            styleRunProperties1.Append(runFonts1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        private static Style GenerateTOC1(Document doc)
        {
            string id = IDGenerator.GenerateStyleID(doc);
            string baseId = Berry.Docx.Documents.ParagraphStyle.CreateBuiltInStyle(BuiltInStyle.Normal, doc).StyleId;

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = id };
            StyleName styleName1 = new StyleName() { Val = "toc 1" };
            BasedOn basedOn1 = new BasedOn() { Val = baseId };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = baseId };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            return style1;
        }

        private static Style GenerateTOC2(Document doc)
        {
            string id = IDGenerator.GenerateStyleID(doc);
            string baseId = Berry.Docx.Documents.ParagraphStyle.CreateBuiltInStyle(BuiltInStyle.Normal, doc).StyleId;

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = id };
            StyleName styleName1 = new StyleName() { Val = "toc 2" };
            BasedOn basedOn1 = new BasedOn() { Val = baseId };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = baseId };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            W.Indentation indentation1 = new W.Indentation() { Left = "405", LeftChars = 200 };

            styleParagraphProperties1.Append(indentation1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            return style1;
        }

        private static Style GenerateTOC3(Document doc)
        {
            string id = IDGenerator.GenerateStyleID(doc);
            string baseId = Berry.Docx.Documents.ParagraphStyle.CreateBuiltInStyle(BuiltInStyle.Normal, doc).StyleId;

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = id };
            StyleName styleName1 = new StyleName() { Val = "toc 3" };
            BasedOn basedOn1 = new BasedOn() { Val = baseId };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = baseId };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            W.Indentation indentation1 = new W.Indentation() { Left = "840", LeftChars = 400 };

            styleParagraphProperties1.Append(indentation1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            return style1;
        }
    }
}
