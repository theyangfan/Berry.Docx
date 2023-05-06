using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;

using SixLabors.ImageSharp;

namespace Berry.Docx
{
    internal class ParagraphItemGenerator
    {
        public static Text GenerateTextRange(string text)
        {
            Run run = new Run();

            //RunProperties rPr = new RunProperties();
            //RunFonts rFonts = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            //rPr.AddChild(rFonts);

            Text text1 = new Text();
            text1.Text = text;

            //run.AddChild(rPr);
            run.AddChild(text1);

            return text1;
        }

        public static Break GenerateBreak()
        {
            Run run = new Run();
            Break br = new Break();
            run.AddChild(br);
            return br;
        }

        public static TabChar GenerateTab()
        {
            Run run = new Run();
            TabChar tab = new TabChar();
            run.AddChild(tab);
            return tab;
        }

        public static NoBreakHyphen GenerateNoBreakHyphen()
        {
            Run run = new Run();
            NoBreakHyphen hyphen = new NoBreakHyphen();
            run.AddChild(hyphen);
            return hyphen;
        }

        public static FieldChar GenerateFieldChar()
        {
            Run run = new Run();
            FieldChar field = new FieldChar();
            run.AddChild(field);
            return field;
        }

        public static FieldCode GenerateFieldCode()
        {
            Run run = new Run();
            FieldCode field = new FieldCode();
            run.AddChild(field);
            return field;
        }

        public static SimpleField GenerateSimpleField()
        {
            SimpleField field = new SimpleField();
            return field;
        }

        public static Run GenerateDrawing(string rId, string filename, float maxWidth, float maxHeight)
        {
            Image image = Image.Load(filename);
            long width = image.Width;
            long height = image.Height;
            image.Dispose();
            float ratio = (float)width / height;
            if (width > maxWidth)
            {
                width = (long)maxWidth;
                height = (long)(maxWidth / ratio);
            }
            if (height > maxHeight)
            {
                height = (long)maxHeight;
                width = (long)(maxHeight * ratio);
            }
            width *= 12700;
            height *= 12700;
            string name = new FileInfo(filename).Name;
            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            runProperties1.Append(runFonts1);

            Drawing drawing1 = new Drawing();

            Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U };
            Wp.Extent extent1 = new Wp.Extent() { Cx = width, Cy = height };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)1U, Name = name };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture1 = new Pic.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = name };
            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Pic.BlipFill blipFill1 = new Pic.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = rId, CompressionState = A.BlipCompressionValues.Print };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            Pic.ShapeProperties shapeProperties1 = new Pic.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = width, Cy = height };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(nonVisualGraphicFrameDrawingProperties1);
            inline1.Append(graphic1);

            drawing1.Append(inline1);

            run1.Append(runProperties1);
            run1.Append(drawing1);
            return run1;
        }

        public static BookmarkStart GenerateBookmarkStart(string id, string name)
        {
            BookmarkStart bookmarkStart = new BookmarkStart();
            bookmarkStart.Id = id;
            bookmarkStart.Name = name;
            return bookmarkStart;
        }

        public static BookmarkEnd GenerateBookmarkEnd(string id)
        {
            BookmarkEnd bookmarkEnd = new BookmarkEnd();
            bookmarkEnd.Id = id;
            return bookmarkEnd;
        }
    }
}
