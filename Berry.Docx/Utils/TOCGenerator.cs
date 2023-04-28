using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace Berry.Docx
{
    internal class TOCGenerator
    {
        public static SdtBlock Generate(int fromLvl, int toLvl)
        {
            SdtBlock sdtBlock1 = new SdtBlock();

            SdtProperties sdtProperties1 = new SdtProperties();

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Color color1 = new Color() { Val = "auto" };
            FontSize fontSize1 = new FontSize() { Val = "21" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };
            Languages languages1 = new Languages() { Val = "zh-CN" };

            runProperties1.Append(runFonts1);
            runProperties1.Append(color1);
            runProperties1.Append(fontSize1);
            runProperties1.Append(fontSizeComplexScript1);
            runProperties1.Append(languages1);
            SdtId sdtId1 = new SdtId() { Val = 632253 };

            SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
            DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Table of Contents" };
            DocPartUnique docPartUnique1 = new DocPartUnique();

            sdtContentDocPartObject1.Append(docPartGallery1);
            sdtContentDocPartObject1.Append(docPartUnique1);

            sdtProperties1.Append(runProperties1);
            //sdtProperties1.Append(sdtId1);
            sdtProperties1.Append(sdtContentDocPartObject1);

            SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

            RunProperties runProperties2 = new RunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();

            runProperties2.Append(bold1);
            runProperties2.Append(boldComplexScript1);

            sdtEndCharProperties1.Append(runProperties2);

            SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

            Paragraph paragraph1 = new Paragraph();

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "TOC1" };
            paragraphProperties2.Append(paragraphStyleId2);

            Run run1 = new Run();
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
            run1.Append(fieldChar1);

            Run run2 = new Run();
            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = $" TOC \\o \"{fromLvl}-{toLvl}\" \\h \\z \\u ";
            run2.Append(fieldCode1);

            Run run3 = new Run();
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };
            run3.Append(fieldChar2);

            Run run4 = new Run();
            Text text = new Text();
            text.Text = "请手动更新目录! ";
            run4.Append(text);

            Run run5 = new Run();
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };
            run5.Append(fieldChar3);

            paragraph1.Append(paragraphProperties2);
            paragraph1.Append(run1);
            paragraph1.Append(run2);
            paragraph1.Append(run3);
            paragraph1.Append(run4);
            paragraph1.Append(run5);

            sdtContentBlock1.Append(paragraph1);

            sdtBlock1.Append(sdtProperties1);
            sdtBlock1.Append(sdtEndCharProperties1);
            sdtBlock1.Append(sdtContentBlock1);
            
            return sdtBlock1;
        }
    }
}
