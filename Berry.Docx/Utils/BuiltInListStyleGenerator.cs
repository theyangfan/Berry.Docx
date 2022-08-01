using System;
using System.Collections.Generic;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx
{
    internal class BuiltInListStyleGenerator
    {
        public static AbstractNum Generate(Document doc, BuiltInListStyle style)
        {
            switch (style)
            {
                case BuiltInListStyle.Style1:
                    return GenerateStyle1(doc);
                case BuiltInListStyle.Style2:
                    return GenerateStyle2(doc);
                case BuiltInListStyle.Style3:
                    return GenerateStyle3(doc);
                case BuiltInListStyle.Style4:
                    return GenerateStyle4(doc);
                default:
                    return null;
            }
        }

        /// <summary>
        /// <para>1 -------</para>
        /// <para>1.1 -----</para>
        /// <para>1.1.1 ---</para>
        /// </summary>
        private static AbstractNum GenerateStyle1(Document doc)
        {
            int id = IDGenerator.GenerateAbstractNumId(doc);

            AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = id };
            abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };

            Level level1 = new Level() { LevelIndex = 0 };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix1 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText1 = new LevelText() { Val = "%1" };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
            W.Indentation indentation1 = new W.Indentation() { Left = "425", Hanging = "425" };

            previousParagraphProperties1.Append(indentation1);

            NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
            RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties1.Append(runFonts1);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelSuffix1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);
            level1.Append(numberingSymbolRunProperties1);

            Level level2 = new Level() { LevelIndex = 1 };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix2 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText2 = new LevelText() { Val = "%1.%2" };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
            W.Indentation indentation2 = new W.Indentation() { Left = "992", Hanging = "567" };

            previousParagraphProperties2.Append(indentation2);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelSuffix2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);

            Level level3 = new Level() { LevelIndex = 2 };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix3 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText3 = new LevelText() { Val = "%1.%2.%3" };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
            W.Indentation indentation3 = new W.Indentation() { Left = "1418", Hanging = "567" };

            previousParagraphProperties3.Append(indentation3);

            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(levelSuffix3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);

            Level level4 = new Level() { LevelIndex = 3 };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix4 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText4 = new LevelText() { Val = "%1.%2.%3.%4" };
            LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
            W.Indentation indentation4 = new W.Indentation() { Left = "1984", Hanging = "708" };

            previousParagraphProperties4.Append(indentation4);

            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            level4.Append(levelSuffix4);
            level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);

            Level level5 = new Level() { LevelIndex = 4 };
            StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix5 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText5 = new LevelText() { Val = "%1.%2.%3.%4.%5" };
            LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
            W.Indentation indentation5 = new W.Indentation() { Left = "2551", Hanging = "850" };

            previousParagraphProperties5.Append(indentation5);

            level5.Append(startNumberingValue5);
            level5.Append(numberingFormat5);
            level5.Append(levelSuffix5);
            level5.Append(levelText5);
            level5.Append(levelJustification5);
            level5.Append(previousParagraphProperties5);

            Level level6 = new Level() { LevelIndex = 5 };
            StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix6 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText6 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6" };
            LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
            W.Indentation indentation6 = new W.Indentation() { Left = "3260", Hanging = "1134" };

            previousParagraphProperties6.Append(indentation6);

            level6.Append(startNumberingValue6);
            level6.Append(numberingFormat6);
            level6.Append(levelSuffix6);
            level6.Append(levelText6);
            level6.Append(levelJustification6);
            level6.Append(previousParagraphProperties6);

            Level level7 = new Level() { LevelIndex = 6 };
            StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix7 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText7 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7" };
            LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
            W.Indentation indentation7 = new W.Indentation() { Left = "3827", Hanging = "1276" };

            previousParagraphProperties7.Append(indentation7);

            level7.Append(startNumberingValue7);
            level7.Append(numberingFormat7);
            level7.Append(levelSuffix7);
            level7.Append(levelText7);
            level7.Append(levelJustification7);
            level7.Append(previousParagraphProperties7);

            Level level8 = new Level() { LevelIndex = 7 };
            StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix8 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText8 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8" };
            LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
            W.Indentation indentation8 = new W.Indentation() { Left = "4394", Hanging = "1418" };

            previousParagraphProperties8.Append(indentation8);

            level8.Append(startNumberingValue8);
            level8.Append(numberingFormat8);
            level8.Append(levelSuffix8);
            level8.Append(levelText8);
            level8.Append(levelJustification8);
            level8.Append(previousParagraphProperties8);

            Level level9 = new Level() { LevelIndex = 8 };
            StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix9 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText9 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8.%9" };
            LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
            W.Indentation indentation9 = new W.Indentation() { Left = "5102", Hanging = "1700" };

            previousParagraphProperties9.Append(indentation9);

            level9.Append(startNumberingValue9);
            level9.Append(numberingFormat9);
            level9.Append(levelSuffix9);
            level9.Append(levelText9);
            level9.Append(levelJustification9);
            level9.Append(previousParagraphProperties9);

            abstractNum1.Append(multiLevelType1);
            abstractNum1.Append(level1);
            abstractNum1.Append(level2);
            abstractNum1.Append(level3);
            abstractNum1.Append(level4);
            abstractNum1.Append(level5);
            abstractNum1.Append(level6);
            abstractNum1.Append(level7);
            abstractNum1.Append(level8);
            abstractNum1.Append(level9);
            return abstractNum1;
        }

        /// <summary>
        /// <para>1. -------</para>
        /// <para>1.1. -----</para>
        /// <para>1.1.1. ---</para>
        /// </summary>
        private static AbstractNum GenerateStyle2(Document doc)
        {
            int id = IDGenerator.GenerateAbstractNumId(doc);

            AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = id };
            abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };

            Level level1 = new Level() { LevelIndex = 0 };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix1 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText1 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
            W.Indentation indentation1 = new W.Indentation() { Left = "425", Hanging = "425" };

            previousParagraphProperties1.Append(indentation1);

            NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
            RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties1.Append(runFonts1);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelSuffix1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);
            level1.Append(numberingSymbolRunProperties1);

            Level level2 = new Level() { LevelIndex = 1 };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix2 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText2 = new LevelText() { Val = "%1.%2." };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
            W.Indentation indentation2 = new W.Indentation() { Left = "567", Hanging = "567" };

            previousParagraphProperties2.Append(indentation2);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelSuffix2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);

            Level level3 = new Level() { LevelIndex = 2 };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix3 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText3 = new LevelText() { Val = "%1.%2.%3." };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
            W.Indentation indentation3 = new W.Indentation() { Left = "709", Hanging = "709" };

            previousParagraphProperties3.Append(indentation3);

            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(levelSuffix3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);

            Level level4 = new Level() { LevelIndex = 3 };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix4 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText4 = new LevelText() { Val = "%1.%2.%3.%4." };
            LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
            W.Indentation indentation4 = new W.Indentation() { Left = "851", Hanging = "851" };

            previousParagraphProperties4.Append(indentation4);

            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            level4.Append(levelSuffix4);
            level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);

            Level level5 = new Level() { LevelIndex = 4 };
            StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix5 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText5 = new LevelText() { Val = "%1.%2.%3.%4.%5." };
            LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
            W.Indentation indentation5 = new W.Indentation() { Left = "992", Hanging = "992" };

            previousParagraphProperties5.Append(indentation5);

            level5.Append(startNumberingValue5);
            level5.Append(numberingFormat5);
            level5.Append(levelSuffix5);
            level5.Append(levelText5);
            level5.Append(levelJustification5);
            level5.Append(previousParagraphProperties5);

            Level level6 = new Level() { LevelIndex = 5 };
            StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix6 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText6 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6." };
            LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
            W.Indentation indentation6 = new W.Indentation() { Left = "1134", Hanging = "1134" };

            previousParagraphProperties6.Append(indentation6);

            level6.Append(startNumberingValue6);
            level6.Append(numberingFormat6);
            level6.Append(levelSuffix6);
            level6.Append(levelText6);
            level6.Append(levelJustification6);
            level6.Append(previousParagraphProperties6);

            Level level7 = new Level() { LevelIndex = 6 };
            StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix7 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText7 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7." };
            LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
            W.Indentation indentation7 = new W.Indentation() { Left = "1276", Hanging = "1276" };

            previousParagraphProperties7.Append(indentation7);

            level7.Append(startNumberingValue7);
            level7.Append(numberingFormat7);
            level7.Append(levelSuffix7);
            level7.Append(levelText7);
            level7.Append(levelJustification7);
            level7.Append(previousParagraphProperties7);

            Level level8 = new Level() { LevelIndex = 7 };
            StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix8 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText8 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8." };
            LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
            W.Indentation indentation8 = new W.Indentation() { Left = "1418", Hanging = "1418" };

            previousParagraphProperties8.Append(indentation8);

            level8.Append(startNumberingValue8);
            level8.Append(numberingFormat8);
            level8.Append(levelSuffix8);
            level8.Append(levelText8);
            level8.Append(levelJustification8);
            level8.Append(previousParagraphProperties8);

            Level level9 = new Level() { LevelIndex = 8 };
            StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix9 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText9 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8.%9." };
            LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
            W.Indentation indentation9 = new W.Indentation() { Left = "1559", Hanging = "1559" };

            previousParagraphProperties9.Append(indentation9);

            level9.Append(startNumberingValue9);
            level9.Append(numberingFormat9);
            level9.Append(levelSuffix9);
            level9.Append(levelText9);
            level9.Append(levelJustification9);
            level9.Append(previousParagraphProperties9);

            abstractNum1.Append(multiLevelType1);
            abstractNum1.Append(level1);
            abstractNum1.Append(level2);
            abstractNum1.Append(level3);
            abstractNum1.Append(level4);
            abstractNum1.Append(level5);
            abstractNum1.Append(level6);
            abstractNum1.Append(level7);
            abstractNum1.Append(level8);
            abstractNum1.Append(level9);
            return abstractNum1;
        }

        /// <summary>
        /// <para>第1章 -------</para>
        /// <para>1.1 -----</para>
        /// <para>1.1.1 ---</para>
        /// </summary>
        private static AbstractNum GenerateStyle3(Document doc)
        {
            int id = IDGenerator.GenerateAbstractNumId(doc);

            AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = id };
            abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };

            Level level1 = new Level() { LevelIndex = 0 };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix1 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText1 = new LevelText() { Val = "第%1章" };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
            W.Indentation indentation1 = new W.Indentation() { Left = "425", Hanging = "425" };

            previousParagraphProperties1.Append(indentation1);

            NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
            RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties1.Append(runFonts1);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelSuffix1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);
            level1.Append(numberingSymbolRunProperties1);

            Level level2 = new Level() { LevelIndex = 1 };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix2 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText2 = new LevelText() { Val = "%1.%2" };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
            W.Indentation indentation2 = new W.Indentation() { Left = "992", Hanging = "567" };

            previousParagraphProperties2.Append(indentation2);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelSuffix2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);

            Level level3 = new Level() { LevelIndex = 2 };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix3 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText3 = new LevelText() { Val = "%1.%2.%3" };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
            W.Indentation indentation3 = new W.Indentation() { Left = "1418", Hanging = "567" };

            previousParagraphProperties3.Append(indentation3);

            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(levelSuffix3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);

            Level level4 = new Level() { LevelIndex = 3 };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix4 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText4 = new LevelText() { Val = "%1.%2.%3.%4" };
            LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
            W.Indentation indentation4 = new W.Indentation() { Left = "1984", Hanging = "708" };

            previousParagraphProperties4.Append(indentation4);

            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            level4.Append(levelSuffix4);
            level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);

            Level level5 = new Level() { LevelIndex = 4 };
            StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix5 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText5 = new LevelText() { Val = "%1.%2.%3.%4.%5" };
            LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
            W.Indentation indentation5 = new W.Indentation() { Left = "2551", Hanging = "850" };

            previousParagraphProperties5.Append(indentation5);

            level5.Append(startNumberingValue5);
            level5.Append(numberingFormat5);
            level5.Append(levelSuffix5);
            level5.Append(levelText5);
            level5.Append(levelJustification5);
            level5.Append(previousParagraphProperties5);

            Level level6 = new Level() { LevelIndex = 5 };
            StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix6 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText6 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6" };
            LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
            W.Indentation indentation6 = new W.Indentation() { Left = "3260", Hanging = "1134" };

            previousParagraphProperties6.Append(indentation6);

            level6.Append(startNumberingValue6);
            level6.Append(numberingFormat6);
            level6.Append(levelSuffix6);
            level6.Append(levelText6);
            level6.Append(levelJustification6);
            level6.Append(previousParagraphProperties6);

            Level level7 = new Level() { LevelIndex = 6 };
            StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix7 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText7 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7" };
            LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
            W.Indentation indentation7 = new W.Indentation() { Left = "3827", Hanging = "1276" };

            previousParagraphProperties7.Append(indentation7);

            level7.Append(startNumberingValue7);
            level7.Append(numberingFormat7);
            level7.Append(levelSuffix7);
            level7.Append(levelText7);
            level7.Append(levelJustification7);
            level7.Append(previousParagraphProperties7);

            Level level8 = new Level() { LevelIndex = 7 };
            StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix8 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText8 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8" };
            LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
            W.Indentation indentation8 = new W.Indentation() { Left = "4394", Hanging = "1418" };

            previousParagraphProperties8.Append(indentation8);

            level8.Append(startNumberingValue8);
            level8.Append(numberingFormat8);
            level8.Append(levelSuffix8);
            level8.Append(levelText8);
            level8.Append(levelJustification8);
            level8.Append(previousParagraphProperties8);

            Level level9 = new Level() { LevelIndex = 8 };
            StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelSuffix levelSuffix9 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText9 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8.%9" };
            LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
            W.Indentation indentation9 = new W.Indentation() { Left = "5102", Hanging = "1700" };

            previousParagraphProperties9.Append(indentation9);

            level9.Append(startNumberingValue9);
            level9.Append(numberingFormat9);
            level9.Append(levelSuffix9);
            level9.Append(levelText9);
            level9.Append(levelJustification9);
            level9.Append(previousParagraphProperties9);

            abstractNum1.Append(multiLevelType1);
            abstractNum1.Append(level1);
            abstractNum1.Append(level2);
            abstractNum1.Append(level3);
            abstractNum1.Append(level4);
            abstractNum1.Append(level5);
            abstractNum1.Append(level6);
            abstractNum1.Append(level7);
            abstractNum1.Append(level8);
            abstractNum1.Append(level9);
            return abstractNum1;
        }

        /// <summary>
        /// <para>一 -------</para>
        /// <para>1.1 -----</para>
        /// <para>1.1.1 ---</para>
        /// </summary>
        private static AbstractNum GenerateStyle4(Document doc)
        {
            int id = IDGenerator.GenerateAbstractNumId(doc);

            AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = id };
            abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };

            Level level1 = new Level() { LevelIndex = 0 };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.ChineseCountingThousand };
            LevelSuffix levelSuffix1 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText1 = new LevelText() { Val = "%1" };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
            W.Indentation indentation1 = new W.Indentation() { Left = "425", Hanging = "425" };

            previousParagraphProperties1.Append(indentation1);

            NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
            RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties1.Append(runFonts1);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelSuffix1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);
            level1.Append(numberingSymbolRunProperties1);

            Level level2 = new Level() { LevelIndex = 1 };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            IsLegalNumberingStyle isLegalNumberingStyle2 = new IsLegalNumberingStyle();
            LevelSuffix levelSuffix2 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText2 = new LevelText() { Val = "%1.%2" };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
            W.Indentation indentation2 = new W.Indentation() { Left = "992", Hanging = "567" };

            previousParagraphProperties2.Append(indentation2);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(isLegalNumberingStyle2);
            level2.Append(levelSuffix2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);

            Level level3 = new Level() { LevelIndex = 2 };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            IsLegalNumberingStyle isLegalNumberingStyle3 = new IsLegalNumberingStyle();
            LevelSuffix levelSuffix3 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText3 = new LevelText() { Val = "%1.%2.%3" };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
            W.Indentation indentation3 = new W.Indentation() { Left = "1418", Hanging = "567" };

            previousParagraphProperties3.Append(indentation3);

            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(isLegalNumberingStyle3);
            level3.Append(levelSuffix3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);

            Level level4 = new Level() { LevelIndex = 3 };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            IsLegalNumberingStyle isLegalNumberingStyle4 = new IsLegalNumberingStyle();
            LevelSuffix levelSuffix4 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText4 = new LevelText() { Val = "%1.%2.%3.%4" };
            LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
            W.Indentation indentation4 = new W.Indentation() { Left = "1984", Hanging = "708" };

            previousParagraphProperties4.Append(indentation4);

            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            level4.Append(isLegalNumberingStyle4);
            level4.Append(levelSuffix4);
            level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);

            Level level5 = new Level() { LevelIndex = 4 };
            StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            IsLegalNumberingStyle isLegalNumberingStyle5 = new IsLegalNumberingStyle();
            LevelSuffix levelSuffix5 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText5 = new LevelText() { Val = "%1.%2.%3.%4.%5" };
            LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
            W.Indentation indentation5 = new W.Indentation() { Left = "2551", Hanging = "850" };

            previousParagraphProperties5.Append(indentation5);

            level5.Append(startNumberingValue5);
            level5.Append(numberingFormat5);
            level5.Append(isLegalNumberingStyle5);
            level5.Append(levelSuffix5);
            level5.Append(levelText5);
            level5.Append(levelJustification5);
            level5.Append(previousParagraphProperties5);

            Level level6 = new Level() { LevelIndex = 5 };
            StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            IsLegalNumberingStyle isLegalNumberingStyle6 = new IsLegalNumberingStyle();
            LevelSuffix levelSuffix6 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText6 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6" };
            LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
            W.Indentation indentation6 = new W.Indentation() { Left = "3260", Hanging = "1134" };

            previousParagraphProperties6.Append(indentation6);

            level6.Append(startNumberingValue6);
            level6.Append(numberingFormat6);
            level6.Append(isLegalNumberingStyle6);
            level6.Append(levelSuffix6);
            level6.Append(levelText6);
            level6.Append(levelJustification6);
            level6.Append(previousParagraphProperties6);

            Level level7 = new Level() { LevelIndex = 6 };
            StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            IsLegalNumberingStyle isLegalNumberingStyle7 = new IsLegalNumberingStyle();
            LevelSuffix levelSuffix7 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText7 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7" };
            LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
            W.Indentation indentation7 = new W.Indentation() { Left = "3827", Hanging = "1276" };

            previousParagraphProperties7.Append(indentation7);

            level7.Append(startNumberingValue7);
            level7.Append(numberingFormat7);
            level7.Append(isLegalNumberingStyle7);
            level7.Append(levelSuffix7);
            level7.Append(levelText7);
            level7.Append(levelJustification7);
            level7.Append(previousParagraphProperties7);

            Level level8 = new Level() { LevelIndex = 7 };
            StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            IsLegalNumberingStyle isLegalNumberingStyle8 = new IsLegalNumberingStyle();
            LevelSuffix levelSuffix8 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText8 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8" };
            LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
            W.Indentation indentation8 = new W.Indentation() { Left = "4394", Hanging = "1418" };

            previousParagraphProperties8.Append(indentation8);

            level8.Append(startNumberingValue8);
            level8.Append(numberingFormat8);
            level8.Append(isLegalNumberingStyle8);
            level8.Append(levelSuffix8);
            level8.Append(levelText8);
            level8.Append(levelJustification8);
            level8.Append(previousParagraphProperties8);

            Level level9 = new Level() { LevelIndex = 8 };
            StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            IsLegalNumberingStyle isLegalNumberingStyle9 = new IsLegalNumberingStyle();
            LevelSuffix levelSuffix9 = new LevelSuffix() { Val = LevelSuffixValues.Space };
            LevelText levelText9 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8.%9" };
            LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
            W.Indentation indentation9 = new W.Indentation() { Left = "5102", Hanging = "1700" };

            previousParagraphProperties9.Append(indentation9);

            level9.Append(startNumberingValue9);
            level9.Append(numberingFormat9);
            level9.Append(isLegalNumberingStyle9);
            level9.Append(levelSuffix9);
            level9.Append(levelText9);
            level9.Append(levelJustification9);
            level9.Append(previousParagraphProperties9);

            abstractNum1.Append(multiLevelType1);
            abstractNum1.Append(level1);
            abstractNum1.Append(level2);
            abstractNum1.Append(level3);
            abstractNum1.Append(level4);
            abstractNum1.Append(level5);
            abstractNum1.Append(level6);
            abstractNum1.Append(level7);
            abstractNum1.Append(level8);
            abstractNum1.Append(level9);
            return abstractNum1;
        }
    }
}
