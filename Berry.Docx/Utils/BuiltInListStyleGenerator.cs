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
                default:
                    return GenerateStyle1(doc);
            }
        }

        private static AbstractNum GenerateStyle1(Document doc)
        {
            int id = IDGenerator.GenerateAbstractNumId(doc);

            AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = id };
            abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };

            Level level1 = new Level() { LevelIndex = 0 };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText1 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
            W.Indentation indentation1 = new W.Indentation() { Left = "420", Hanging = "420" };

            previousParagraphProperties1.Append(indentation1);

            NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
            RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties1.Append(runFonts1);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);
            level1.Append(numberingSymbolRunProperties1);

            Level level2 = new Level() { LevelIndex = 1, Tentative = true };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText2 = new LevelText() { Val = "%2)" };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
            W.Indentation indentation2 = new W.Indentation() { Left = "840", Hanging = "420" };

            previousParagraphProperties2.Append(indentation2);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);

            Level level3 = new Level() { LevelIndex = 2, Tentative = true };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText3 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
            W.Indentation indentation3 = new W.Indentation() { Left = "1260", Hanging = "420" };

            previousParagraphProperties3.Append(indentation3);

            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);

            Level level4 = new Level() { LevelIndex = 3, Tentative = true };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText4 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
            W.Indentation indentation4 = new W.Indentation() { Left = "1680", Hanging = "420" };

            previousParagraphProperties4.Append(indentation4);

            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);

            Level level5 = new Level() { LevelIndex = 4, Tentative = true };
            StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText5 = new LevelText() { Val = "%5)" };
            LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
            W.Indentation indentation5 = new W.Indentation() { Left = "2100", Hanging = "420" };

            previousParagraphProperties5.Append(indentation5);

            level5.Append(startNumberingValue5);
            level5.Append(numberingFormat5);
            level5.Append(levelText5);
            level5.Append(levelJustification5);
            level5.Append(previousParagraphProperties5);

            Level level6 = new Level() { LevelIndex = 5, Tentative = true };
            StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText6 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
            W.Indentation indentation6 = new W.Indentation() { Left = "2520", Hanging = "420" };

            previousParagraphProperties6.Append(indentation6);

            level6.Append(startNumberingValue6);
            level6.Append(numberingFormat6);
            level6.Append(levelText6);
            level6.Append(levelJustification6);
            level6.Append(previousParagraphProperties6);

            Level level7 = new Level() { LevelIndex = 6, Tentative = true };
            StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText7 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
            W.Indentation indentation7 = new W.Indentation() { Left = "2940", Hanging = "420" };

            previousParagraphProperties7.Append(indentation7);

            level7.Append(startNumberingValue7);
            level7.Append(numberingFormat7);
            level7.Append(levelText7);
            level7.Append(levelJustification7);
            level7.Append(previousParagraphProperties7);

            Level level8 = new Level() { LevelIndex = 7, Tentative = true };
            StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText8 = new LevelText() { Val = "%8)" };
            LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
            W.Indentation indentation8 = new W.Indentation() { Left = "3360", Hanging = "420" };

            previousParagraphProperties8.Append(indentation8);

            level8.Append(startNumberingValue8);
            level8.Append(numberingFormat8);
            level8.Append(levelText8);
            level8.Append(levelJustification8);
            level8.Append(previousParagraphProperties8);

            Level level9 = new Level() { LevelIndex = 8, Tentative = true };
            StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText9 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
            W.Indentation indentation9 = new W.Indentation() { Left = "3780", Hanging = "420" };

            previousParagraphProperties9.Append(indentation9);

            level9.Append(startNumberingValue9);
            level9.Append(numberingFormat9);
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
