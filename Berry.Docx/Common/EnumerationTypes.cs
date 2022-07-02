using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Defines the DocumentObjectType enumeration.
    /// </summary>
    public enum DocumentObjectType
    {
        /// <summary>
        /// Invalid Type.
        /// </summary>
        Invalid = -1,
        /// <summary>
        /// Paragraph.
        /// </summary>
        Paragraph,
        /// <summary>
        /// Text range.
        /// </summary>
        TextRange,
        /// <summary>
        /// Table.
        /// </summary>
        Table,
        /// <summary>
        /// Table row.
        /// </summary>
        TableRow,
        /// <summary>
        /// Table cell.
        /// </summary>
        TableCell,
        /// <summary>
        /// Section.
        /// </summary>
        Section,
        /// <summary>
        /// Body range.
        /// </summary>
        BodyRange,
        /// <summary>
        /// Picture.
        /// </summary>
        Picture,
        /// <summary>
        /// Shape.
        /// </summary>
        Shape,
        /// <summary>
        /// Shape group.
        /// </summary>
        GroupShape,
        /// <summary>
        /// Canvas.
        /// </summary>
        Canvas,
        /// <summary>
        /// Diagram.
        /// </summary>
        Diagram,
        /// <summary>
        /// Chart.
        /// </summary>
        Chart,
        /// <summary>
        /// FootnoteReference.
        /// </summary>
        FootnoteReference,
        /// <summary>
        /// EndnoteReference.
        /// </summary>
        EndnoteReference,
        /// <summary>
        /// Embedded Object.
        /// </summary>
        EmbeddedObject,
        /// <summary>
        /// Office Mathematical Text.
        /// </summary>
        OfficeMath,
        SdtBlock,
        SdtContent,
        SdtProperties,
        Break
    }

    /// <summary>
    /// Defines the LineSpacingRule enumeration.
    /// </summary>
    public enum LineSpacingRule
    {
        /// <summary>
        /// Minimum Line Height.
        /// </summary>
        AtLeast = 0,
        /// <summary>
        /// Exact Line Height.
        /// </summary>
        Exactly = 1,
        /// <summary>
        /// The line spacing is specified in the LineSpacing property as the number of lines.
        /// One line equals 12 points.
        /// </summary>
        Multiple = 2
    }

    /// <summary>
    /// Defines the MultiPage enumeration.
    /// </summary>
    public enum MultiPage
    {
        /// <summary>
        /// Normal page number range.
        /// </summary>
        Normal = 0,
        /// <summary>
        /// MirrorMargins page number range.
        /// </summary>
        MirrorMargins = 1,
        /// <summary>
        /// PrintTwoOnOne page number range.
        /// </summary>
        PrintTwoOnOne = 2
    }
    /// <summary>
    /// Defines the Text flow Direction enumeration.
    /// </summary>
    public enum TextFlowDirection
    {
        /// <summary>
        /// Lef to Right, Top to Bottom.
        /// </summary>
        Horizontal = 0,
        /// <summary>
        /// Top to Bottom, Right to Left.
        /// </summary>
        Vertical = 1,
        /// <summary>
        /// Left to Right, Top to Bottom Rotated.
        /// </summary>
        RotateAsianChars270 = 2
    }
    /// <summary>
    /// Defines the DocGridType enumeration.
    /// </summary>
    public enum DocGridType
    {
        /// <summary>
        /// No Document Grid.
        /// </summary>
        None = 0,
        /// <summary>
        /// Line Grid Only.
        /// </summary>
        Lines = 1,
        /// <summary>
        /// Line and Character Grid.
        /// </summary>
        LinesAndChars = 2,
        /// <summary>
        /// Snap to Character Grid.
        /// </summary>
        SnapToChars = 3
    }

    /// <summary>
    /// Defines the StyleType enumeration.
    /// </summary>
    public enum StyleType
    {
        /// <summary>
        /// Paragraph Style.
        /// </summary>
        Paragraph = 0,
        /// <summary>
        /// Character Style.
        /// </summary>
        Character = 1,
        /// <summary>
        /// Table Style.
        /// </summary>
        Table = 2,
        /// <summary>
        /// Numbering Style.
        /// </summary>
        Numbering = 3
    }

    /// <summary>
    /// Defines the JustificationType enumeration.
    /// </summary>
    public enum JustificationType
    {
        /// <summary>
        /// Align Left.
        /// </summary>
        [Description("左对齐")]
        Left = 0,
        /// <summary>
        /// Align Center.
        /// </summary>
        [Description("居中")]
        Center = 1,
        /// <summary>
        /// Align Right.
        /// </summary>
        [Description("右对齐")]
        Right = 2,
        /// <summary>
        /// Justified.
        /// </summary>
        [Description("两端对齐")]
        Both = 3,
        /// <summary>
        /// Distribute All Characters Equally.
        /// </summary>
        [Description("分散对齐")]
        Distribute = 4
    }

    /// <summary>
    /// Defines the OutlineLevelType enumeration.
    /// </summary>
    public enum OutlineLevelType
    {
        /// <summary>
        /// Level 1.
        /// </summary>
        [Description("1 级")]
        Level1 = 0,
        /// <summary>
        /// Level 2.
        /// </summary>
        /// 
        [Description("2 级")]
        Level2 = 1,
        /// <summary>
        /// Level 3.
        /// </summary>
        [Description("3 级")]
        Level3 = 2,
        /// <summary>
        /// Level 4.
        /// </summary>
        [Description("4 级")]
        Level4 = 3,
        /// <summary>
        /// Level 5.
        /// </summary>
        [Description("5 级")]
        Level5 = 4,
        /// <summary>
        /// Level 6.
        /// </summary>
        [Description("6 级")]
        Level6 = 5,
        /// <summary>
        /// Level 7.
        /// </summary>
        [Description("7 级")]
        Level7 = 6,
        /// <summary>
        /// Level 8.
        /// </summary>
        [Description("8 级")]
        Level8 = 7,
        /// <summary>
        /// Level 9.
        /// </summary>
        [Description("9 级")]
        Level9 = 8,
        /// <summary>
        /// Body Text.
        /// </summary>
        [Description("正文文本")]
        BodyText = 9
    }

    /// <summary>
    /// Defines the SectionBreakType enumeration.
    /// </summary>
    public enum SectionBreakType
    {
        /// <summary>
        /// Next Page Section Break.
        /// </summary>
        NextPage = 0,
        /// <summary>
        /// Continuous Section Break.
        /// </summary>
        Continuous = 1,
        /// <summary>
        /// Odd Page Section Break.
        /// </summary>
        OddPage = 2,
        /// <summary>
        /// Even Page Section Break.
        /// </summary>
        EvenPage = 3
    }

    /// <summary>
    /// Defines the PageOrientation enumeration.
    /// </summary>
    public enum PageOrientation
    {
        /// <summary>
        /// Portrait Mode.
        /// </summary>
        [Description("纵向")]
        Portrait = 0,
        /// <summary>
        /// Landscape Mode.
        /// </summary>
        [Description("横向")]
        Landscape = 1
    }

    /// <summary>
    /// Defines the page gutter location enumeration.
    /// </summary>
    public enum GutterLocation
    {
        /// <summary>
        /// Position Gutter At Left of Page.
        /// </summary>
        [Description("靠左")]
        Left = 0,
        /// <summary>
        /// Position Gutter At Top of Page.
        /// </summary>
        [Description("靠上")]
        Top = 1
    }

    /// <summary>
    /// Defines the footnote and endnote numbering restart rule enumeration. 
    /// </summary>
    public enum FootEndnoteNumberRestartRule
    {
        /// <summary>
        /// Continue Numbering From Previous Section.
        /// </summary>
        [Description("连续")]
        Continuous = 0,
        /// <summary>
        /// Restart Numbering For Each Section.
        /// </summary>
        [Description("每节重新编号")]
        EachSection = 1,
        /// <summary>
        /// Restart Numbering On Each Page.
        /// </summary>
        [Description("每页重新编号")]
        EachPage = 2
    }

    /// <summary>
    /// Defines picture text wrapping style enumeration.
    /// </summary>
    public enum TextWrappingStyle
    {
        Inline = 0,
        Floating = 1
    }
    /// <summary>
    /// Defines OLE object type enumeration.
    /// </summary>
    public enum OleObjectType
    {
        Embed = 0,
        Link = 1
    }

    /// <summary>
    /// Defines office math justification type enumeration.
    /// </summary>
    public enum OfficeMathJustificationType
    {
        /// <summary>
        /// Invalid Justification.
        /// </summary>
        Invalid = -1,
        /// <summary>
        /// Left Justification.
        /// </summary>
        [Description("左对齐")]
        Left = 0,
        /// <summary>
        /// Right Justification. 
        /// </summary>
        [Description("右对齐")]
        Right = 1,
        /// <summary>
        /// Center Justification. 
        /// </summary>
        [Description("居中")]
        Center = 2,
        /// <summary>
        /// Center as Group Justification.
        /// </summary>
        [Description("整体居中")]
        CenterGroup = 3
    }

    /// <summary>
    /// Defines the vertical text alignment on page enumeration.
    /// </summary>
    public enum VerticalJustificationType
    {
        /// <summary>
        /// Align Top. 
        /// </summary>
        Top = 0,
        /// <summary>
        /// Align Center.
        /// </summary>
        Center = 1,
        /// <summary>
        /// Vertical Justification.
        /// </summary>
        Both = 2,
        /// <summary>
        /// Align Bottom.
        /// </summary>
        Bottom = 3
    }

    public enum BuiltInStyle
    {
        None = -1,
        /// <summary>
        /// 正文(Normal)
        /// </summary>
        Normal = 0,
        /// <summary>
        /// 标题 1(heading 1)
        /// </summary>
        Heading1 = 1,
        Heading2 = 2,
        Heading3 = 3,
        Heading4 = 4,
        Heading5 = 5,
        Heading6 = 6,
        Heading7 = 7,
        Heading8 = 8,
        Heading9 = 9,
        /// <summary>
        /// 标题(Title)
        /// </summary>
        Title = 10,
        /// <summary>
        /// 副标题(Subtitle)
        /// </summary>
        SubTitle = 11,
        /// <summary>
        /// 目录 1(toc 1)
        /// </summary>
        TOC1 = 12,
        TOC2 = 13,
        TOC3 = 14,
        TOC4 = 15,
        TOC5 = 16,
        TOC6 = 17,
        TOC7 = 18,
        TOC8 = 19,
        TOC9 = 20,
        /// <summary>
        /// 页眉(header)
        /// </summary>
        Header = 16,
        /// <summary>
        /// 页脚(footer)
        /// </summary>
        Footer = 17

    }

    /// <summary>
    /// Defines the BreakType enumeration. 
    /// </summary>
    public enum BreakType
    {
        /// <summary>
        /// Page Break.
        /// </summary>
        Page = 0,
        /// <summary>
        /// Column Break.
        /// </summary>
        Column = 1,
        /// <summary>
        /// Line Break.
        /// </summary>
        TextWrapping = 2
    }

    /// <summary>
    /// Defines the BreakTextRestartLocation enumeration. 
    /// </summary>
    public enum BreakTextRestartLocation
    {
        /// <summary>
        /// Restart On Next Line.
        /// </summary>
        None = 0,
        /// <summary>
        /// Restart In Next Text Region When In Leftmost Position. 
        /// </summary>
        Left = 1,
        /// <summary>
        /// Restart In Next Text Region When In Rightmost Position. 
        /// </summary>
        Right = 2,
        /// <summary>
        /// Restart On Next Full Line. 
        /// </summary>
        All = 3
    }

    /// <summary>
    /// Defines the Subscript/Superscript Value enumeration. 
    /// </summary>
    public enum SubSuperScript
    {
        /// <summary>
        /// Regular Vertical Positioning.
        /// </summary>
        None = 0,
        /// <summary>
        /// Superscript.
        /// </summary>
        SuperScript = 1,
        /// <summary>
        /// Subscript.
        /// </summary>
        SubScript = 2
    }

    /// <summary>
    /// Defines the Underline Style enumeration. 
    /// </summary>
    public enum UnderlineStyle
    {
        /// <summary>
        /// Single Underline.
        /// </summary>
        Single = 0,
        /// <summary>
        /// Underline Non-Space Characters Only.
        /// </summary>
        Words = 1,
        /// <summary>
        /// Double Underline.
        /// </summary>
        Double = 2,
        /// <summary>
        /// Thick Underline.
        /// </summary>
        Thick = 3,
        /// <summary>
        /// Dotted Underline.
        /// </summary>
        Dotted = 4,
        /// <summary>
        /// Thick Dotted Underline.
        /// </summary>
        DottedHeavy = 5,
        /// <summary>
        /// Dashed Underline.
        /// </summary>
        Dash = 6,
        /// <summary>
        /// Thick Dashed Underline.
        /// </summary>
        DashedHeavy = 7,
        /// <summary>
        /// Long Dashed Underline.
        /// </summary>
        DashLong = 8,
        /// <summary>
        /// Thick Long Dashed Underline.
        /// </summary>
        DashLongHeavy = 9,
        /// <summary>
        /// Dash-Dot Underline.
        /// </summary>
        DotDash = 10,
        /// <summary>
        /// Thick Dash-Dot Underline.
        /// </summary>
        DashDotHeavy = 11,
        /// <summary>
        /// Dash-Dot-Dot Underline.
        /// </summary>
        DotDotDash = 12,
        /// <summary>
        /// Thick Dash-Dot-Dot Underline.
        /// </summary>
        DashDotDotHeavy = 13,
        /// <summary>
        /// Wave Underline.
        /// </summary>
        Wave = 14,
        /// <summary>
        /// Heavy Wave Underline.
        /// </summary>
        WavyHeavy = 15,
        /// <summary>
        /// Double Wave Underline.
        /// </summary>
        WavyDouble = 16,
        /// <summary>
        /// No Underline.
        /// </summary>
        None = 17
    }

    /// <summary>
    /// Defines the Indentation Unit enumeration. 
    /// </summary>
    public enum IndentationUnit
    {
        /// <summary>
        /// Character Unit.
        /// </summary>
        Character = 0,
        /// <summary>
        /// Point Unit.
        /// </summary>
        Point = 1
    }

    /// <summary>
    /// Defines the SpecialIndentation Type enumeration. 
    /// </summary>
    public enum SpecialIndentationType
    {
        /// <summary>
        /// None Indentation.
        /// </summary>
        None = 0,
        /// <summary>
        /// FirstLine Indentation.
        /// </summary>
        FirstLine = 1,
        /// <summary>
        /// Hanging Indentation.
        /// </summary>
        Hanging = 2
    }

    /// <summary>
    /// Defines the Spacing Unit enumeration. 
    /// </summary>
    public enum SpacingUnit
    {
        /// <summary>
        /// Line Unit.
        /// </summary>
        Line = 0,
        /// <summary>
        /// Point Unit.
        /// </summary>
        Point = 1
    }

    /// <summary>
    /// Defines the font content type enumeration. 
    /// </summary>
    public enum FontContentType
    {
        /// <summary>
        /// High ANSI Font.
        /// </summary>
        Default = 0,
        /// <summary>
        /// East Asian Font.
        /// </summary>
        EastAsia = 1,
        /// <summary>
        /// Complex Script Font.
        /// </summary>
        ComplexScript = 2
    }
    /// <summary>
    /// Defines the vertical text alignment enumeration. 
    /// </summary>
    public enum VerticalTextAlignment
    {
        /// <summary>
        /// Align Text at Top. 
        /// </summary>
        Top = 0,
        /// <summary>
        /// Align Text at Center. 
        /// </summary>
        Center = 1,
        /// <summary>
        /// Align Text at Baseline. 
        /// </summary>
        Baseline = 2,
        /// <summary>
        /// Align Text at Bottom. 
        /// </summary>
        Bottom = 3,
        /// <summary>
        /// Automatically Determine Alignment. 
        /// </summary>
        Auto = 4
    }

    /// <summary>
    /// Defines the border style enumeration.
    /// </summary>
    public enum BorderStyle
    {
        /// <summary>
        /// No Border.
        /// </summary>
        Nil = 0,
        /// <summary>
        /// No Border.
        /// </summary>
        None = 1,
        /// <summary>
        /// Single Line Border.
        /// </summary>
        Single = 2,
        /// <summary>
        /// Single Line Border.
        /// </summary>
        Thick = 3,
        //
        // 摘要:
        //     Double Line Border.
        //     When the item is serialized out as xml, its value is "double".
        Double = 4,
        //
        // 摘要:
        //     Dotted Line Border.
        //     When the item is serialized out as xml, its value is "dotted".
        Dotted = 5,
        //
        // 摘要:
        //     Dashed Line Border.
        //     When the item is serialized out as xml, its value is "dashed".
        Dashed = 6,
        //
        // 摘要:
        //     Dot Dash Line Border.
        //     When the item is serialized out as xml, its value is "dotDash".
        DotDash = 7,
        //
        // 摘要:
        //     Dot Dot Dash Line Border.
        //     When the item is serialized out as xml, its value is "dotDotDash".
        DotDotDash = 8,
        //
        // 摘要:
        //     Triple Line Border.
        //     When the item is serialized out as xml, its value is "triple".
        Triple = 9,
        //
        // 摘要:
        //     Thin, Thick Line Border.
        //     When the item is serialized out as xml, its value is "thinThickSmallGap".
        ThinThickSmallGap = 10,
        //
        // 摘要:
        //     Thick, Thin Line Border.
        //     When the item is serialized out as xml, its value is "thickThinSmallGap".
        ThickThinSmallGap = 11,
        //
        // 摘要:
        //     Thin, Thick, Thin Line Border.
        //     When the item is serialized out as xml, its value is "thinThickThinSmallGap".
        ThinThickThinSmallGap = 12,
        //
        // 摘要:
        //     Thin, Thick Line Border.
        //     When the item is serialized out as xml, its value is "thinThickMediumGap".
        ThinThickMediumGap = 13,
        //
        // 摘要:
        //     Thick, Thin Line Border.
        //     When the item is serialized out as xml, its value is "thickThinMediumGap".
        ThickThinMediumGap = 14,
        //
        // 摘要:
        //     Thin, Thick, Thin Line Border.
        //     When the item is serialized out as xml, its value is "thinThickThinMediumGap".
        ThinThickThinMediumGap = 15,
        //
        // 摘要:
        //     Thin, Thick Line Border.
        //     When the item is serialized out as xml, its value is "thinThickLargeGap".
        ThinThickLargeGap = 16,
        //
        // 摘要:
        //     Thick, Thin Line Border.
        //     When the item is serialized out as xml, its value is "thickThinLargeGap".
        ThickThinLargeGap = 17,
        //
        // 摘要:
        //     Thin, Thick, Thin Line Border.
        //     When the item is serialized out as xml, its value is "thinThickThinLargeGap".
        ThinThickThinLargeGap = 18,
        //
        // 摘要:
        //     Wavy Line Border.
        //     When the item is serialized out as xml, its value is "wave".
        Wave = 19,
        //
        // 摘要:
        //     Double Wave Line Border.
        //     When the item is serialized out as xml, its value is "doubleWave".
        DoubleWave = 20,
        //
        // 摘要:
        //     Dashed Line Border.
        //     When the item is serialized out as xml, its value is "dashSmallGap".
        DashSmallGap = 21,
        //
        // 摘要:
        //     Dash Dot Strokes Line Border.
        //     When the item is serialized out as xml, its value is "dashDotStroked".
        DashDotStroked = 22,
        //
        // 摘要:
        //     3D Embossed Line Border.
        //     When the item is serialized out as xml, its value is "threeDEmboss".
        ThreeDEmboss = 23,
        //
        // 摘要:
        //     3D Engraved Line Border.
        //     When the item is serialized out as xml, its value is "threeDEngrave".
        ThreeDEngrave = 24,
        //
        // 摘要:
        //     Outset Line Border.
        //     When the item is serialized out as xml, its value is "outset".
        Outset = 25,
        //
        // 摘要:
        //     Inset Line Border.
        //     When the item is serialized out as xml, its value is "inset".
        Inset = 26,
        //
        // 摘要:
        //     Apples Art Border.
        //     When the item is serialized out as xml, its value is "apples".
        Apples = 27,
        //
        // 摘要:
        //     Arched Scallops Art Border.
        //     When the item is serialized out as xml, its value is "archedScallops".
        ArchedScallops = 28,
        //
        // 摘要:
        //     Baby Pacifier Art Border.
        //     When the item is serialized out as xml, its value is "babyPacifier".
        BabyPacifier = 29,
        //
        // 摘要:
        //     Baby Rattle Art Border.
        //     When the item is serialized out as xml, its value is "babyRattle".
        BabyRattle = 30,
        //
        // 摘要:
        //     Three Color Balloons Art Border.
        //     When the item is serialized out as xml, its value is "balloons3Colors".
        Balloons3Colors = 31,
        //
        // 摘要:
        //     Hot Air Balloons Art Border.
        //     When the item is serialized out as xml, its value is "balloonsHotAir".
        BalloonsHotAir = 32,
        //
        // 摘要:
        //     Black Dash Art Border.
        //     When the item is serialized out as xml, its value is "basicBlackDashes".
        BasicBlackDashes = 33,
        //
        // 摘要:
        //     Black Dot Art Border.
        //     When the item is serialized out as xml, its value is "basicBlackDots".
        BasicBlackDots = 34,
        //
        // 摘要:
        //     Black Square Art Border.
        //     When the item is serialized out as xml, its value is "basicBlackSquares".
        BasicBlackSquares = 35,
        //
        // 摘要:
        //     Thin Line Art Border.
        //     When the item is serialized out as xml, its value is "basicThinLines".
        BasicThinLines = 36,
        //
        // 摘要:
        //     White Dash Art Border.
        //     When the item is serialized out as xml, its value is "basicWhiteDashes".
        BasicWhiteDashes = 37,
        //
        // 摘要:
        //     White Dot Art Border.
        //     When the item is serialized out as xml, its value is "basicWhiteDots".
        BasicWhiteDots = 38,
        //
        // 摘要:
        //     White Square Art Border.
        //     When the item is serialized out as xml, its value is "basicWhiteSquares".
        BasicWhiteSquares = 39,
        //
        // 摘要:
        //     Wide Inline Art Border.
        //     When the item is serialized out as xml, its value is "basicWideInline".
        BasicWideInline = 40,
        //
        // 摘要:
        //     Wide Midline Art Border.
        //     When the item is serialized out as xml, its value is "basicWideMidline".
        BasicWideMidline = 41,
        //
        // 摘要:
        //     Wide Outline Art Border.
        //     When the item is serialized out as xml, its value is "basicWideOutline".
        BasicWideOutline = 42,
        //
        // 摘要:
        //     Bats Art Border.
        //     When the item is serialized out as xml, its value is "bats".
        Bats = 43,
        //
        // 摘要:
        //     Birds Art Border.
        //     When the item is serialized out as xml, its value is "birds".
        Birds = 44,
        //
        // 摘要:
        //     Birds Flying Art Border.
        //     When the item is serialized out as xml, its value is "birdsFlight".
        BirdsFlight = 45,
        //
        // 摘要:
        //     Cabin Art Border.
        //     When the item is serialized out as xml, its value is "cabins".
        Cabins = 46,
        //
        // 摘要:
        //     Cake Art Border.
        //     When the item is serialized out as xml, its value is "cakeSlice".
        CakeSlice = 47,
        //
        // 摘要:
        //     Candy Corn Art Border.
        //     When the item is serialized out as xml, its value is "candyCorn".
        CandyCorn = 48,
        //
        // 摘要:
        //     Knot Work Art Border.
        //     When the item is serialized out as xml, its value is "celticKnotwork".
        CelticKnotwork = 49,
        //
        // 摘要:
        //     Certificate Banner Art Border.
        //     When the item is serialized out as xml, its value is "certificateBanner".
        CertificateBanner = 50,
        //
        // 摘要:
        //     Chain Link Art Border.
        //     When the item is serialized out as xml, its value is "chainLink".
        ChainLink = 51,
        //
        // 摘要:
        //     Champagne Bottle Art Border.
        //     When the item is serialized out as xml, its value is "champagneBottle".
        ChampagneBottle = 52,
        //
        // 摘要:
        //     Black and White Bar Art Border.
        //     When the item is serialized out as xml, its value is "checkedBarBlack".
        CheckedBarBlack = 53,
        //
        // 摘要:
        //     Color Checked Bar Art Border.
        //     When the item is serialized out as xml, its value is "checkedBarColor".
        CheckedBarColor = 54,
        //
        // 摘要:
        //     Checkerboard Art Border.
        //     When the item is serialized out as xml, its value is "checkered".
        Checkered = 55,
        //
        // 摘要:
        //     Christmas Tree Art Border.
        //     When the item is serialized out as xml, its value is "christmasTree".
        ChristmasTree = 56,
        //
        // 摘要:
        //     Circles And Lines Art Border.
        //     When the item is serialized out as xml, its value is "circlesLines".
        CirclesLines = 57,
        //
        // 摘要:
        //     Circles and Rectangles Art Border.
        //     When the item is serialized out as xml, its value is "circlesRectangles".
        CirclesRectangles = 58,
        //
        // 摘要:
        //     Wave Art Border.
        //     When the item is serialized out as xml, its value is "classicalWave".
        ClassicalWave = 59,
        //
        // 摘要:
        //     Clocks Art Border.
        //     When the item is serialized out as xml, its value is "clocks".
        Clocks = 60,
        //
        // 摘要:
        //     Compass Art Border.
        //     When the item is serialized out as xml, its value is "compass".
        Compass = 61,
        //
        // 摘要:
        //     Confetti Art Border.
        //     When the item is serialized out as xml, its value is "confetti".
        Confetti = 62,
        //
        // 摘要:
        //     Confetti Art Border.
        //     When the item is serialized out as xml, its value is "confettiGrays".
        ConfettiGrays = 63,
        //
        // 摘要:
        //     Confetti Art Border.
        //     When the item is serialized out as xml, its value is "confettiOutline".
        ConfettiOutline = 64,
        //
        // 摘要:
        //     Confetti Streamers Art Border.
        //     When the item is serialized out as xml, its value is "confettiStreamers".
        ConfettiStreamers = 65,
        //
        // 摘要:
        //     Confetti Art Border.
        //     When the item is serialized out as xml, its value is "confettiWhite".
        ConfettiWhite = 66,
        //
        // 摘要:
        //     Corner Triangle Art Border.
        //     When the item is serialized out as xml, its value is "cornerTriangles".
        CornerTriangles = 67,
        //
        // 摘要:
        //     Dashed Line Art Border.
        //     When the item is serialized out as xml, its value is "couponCutoutDashes".
        CouponCutoutDashes = 68,
        //
        // 摘要:
        //     Dotted Line Art Border.
        //     When the item is serialized out as xml, its value is "couponCutoutDots".
        CouponCutoutDots = 69,
        //
        // 摘要:
        //     Maze Art Border.
        //     When the item is serialized out as xml, its value is "crazyMaze".
        CrazyMaze = 70,
        //
        // 摘要:
        //     Butterfly Art Border.
        //     When the item is serialized out as xml, its value is "creaturesButterfly".
        CreaturesButterfly = 71,
        //
        // 摘要:
        //     Fish Art Border.
        //     When the item is serialized out as xml, its value is "creaturesFish".
        CreaturesFish = 72,
        //
        // 摘要:
        //     Insects Art Border.
        //     When the item is serialized out as xml, its value is "creaturesInsects".
        CreaturesInsects = 73,
        //
        // 摘要:
        //     Ladybug Art Border.
        //     When the item is serialized out as xml, its value is "creaturesLadyBug".
        CreaturesLadyBug = 74,
        //
        // 摘要:
        //     Cross-stitch Art Border.
        //     When the item is serialized out as xml, its value is "crossStitch".
        CrossStitch = 75,
        //
        // 摘要:
        //     Cupid Art Border.
        //     When the item is serialized out as xml, its value is "cup".
        Cup = 76,
        //
        // 摘要:
        //     Archway Art Border.
        //     When the item is serialized out as xml, its value is "decoArch".
        DecoArch = 77,
        //
        // 摘要:
        //     Color Archway Art Border.
        //     When the item is serialized out as xml, its value is "decoArchColor".
        DecoArchColor = 78,
        //
        // 摘要:
        //     Blocks Art Border.
        //     When the item is serialized out as xml, its value is "decoBlocks".
        DecoBlocks = 79,
        //
        // 摘要:
        //     Gray Diamond Art Border.
        //     When the item is serialized out as xml, its value is "diamondsGray".
        DiamondsGray = 80,
        //
        // 摘要:
        //     Double D Art Border.
        //     When the item is serialized out as xml, its value is "doubleD".
        DoubleD = 81,
        //
        // 摘要:
        //     Diamond Art Border.
        //     When the item is serialized out as xml, its value is "doubleDiamonds".
        DoubleDiamonds = 82,
        //
        // 摘要:
        //     Earth Art Border.
        //     When the item is serialized out as xml, its value is "earth1".
        Earth1 = 83,
        //
        // 摘要:
        //     Earth Art Border.
        //     When the item is serialized out as xml, its value is "earth2".
        Earth2 = 84,
        //
        // 摘要:
        //     Shadowed Square Art Border.
        //     When the item is serialized out as xml, its value is "eclipsingSquares1".
        EclipsingSquares1 = 85,
        //
        // 摘要:
        //     Shadowed Square Art Border.
        //     When the item is serialized out as xml, its value is "eclipsingSquares2".
        EclipsingSquares2 = 86,
        //
        // 摘要:
        //     Painted Egg Art Border.
        //     When the item is serialized out as xml, its value is "eggsBlack".
        EggsBlack = 87,
        //
        // 摘要:
        //     Fans Art Border.
        //     When the item is serialized out as xml, its value is "fans".
        Fans = 88,
        //
        // 摘要:
        //     Film Reel Art Border.
        //     When the item is serialized out as xml, its value is "film".
        Film = 89,
        //
        // 摘要:
        //     Firecracker Art Border.
        //     When the item is serialized out as xml, its value is "firecrackers".
        Firecrackers = 90,
        //
        // 摘要:
        //     Flowers Art Border.
        //     When the item is serialized out as xml, its value is "flowersBlockPrint".
        FlowersBlockPrint = 91,
        //
        // 摘要:
        //     Daisy Art Border.
        //     When the item is serialized out as xml, its value is "flowersDaisies".
        FlowersDaisies = 92,
        //
        // 摘要:
        //     Flowers Art Border.
        //     When the item is serialized out as xml, its value is "flowersModern1".
        FlowersModern1 = 93,
        //
        // 摘要:
        //     Flowers Art Border.
        //     When the item is serialized out as xml, its value is "flowersModern2".
        FlowersModern2 = 94,
        //
        // 摘要:
        //     Pansy Art Border.
        //     When the item is serialized out as xml, its value is "flowersPansy".
        FlowersPansy = 95,
        //
        // 摘要:
        //     Red Rose Art Border.
        //     When the item is serialized out as xml, its value is "flowersRedRose".
        FlowersRedRose = 96,
        //
        // 摘要:
        //     Roses Art Border.
        //     When the item is serialized out as xml, its value is "flowersRoses".
        FlowersRoses = 97,
        //
        // 摘要:
        //     Flowers in a Teacup Art Border.
        //     When the item is serialized out as xml, its value is "flowersTeacup".
        FlowersTeacup = 98,
        //
        // 摘要:
        //     Small Flower Art Border.
        //     When the item is serialized out as xml, its value is "flowersTiny".
        FlowersTiny = 99,
        //
        // 摘要:
        //     Gems Art Border.
        //     When the item is serialized out as xml, its value is "gems".
        Gems = 100,
        //
        // 摘要:
        //     Gingerbread Man Art Border.
        //     When the item is serialized out as xml, its value is "gingerbreadMan".
        GingerbreadMan = 101,
        //
        // 摘要:
        //     Triangle Gradient Art Border.
        //     When the item is serialized out as xml, its value is "gradient".
        Gradient = 102,
        //
        // 摘要:
        //     Handmade Art Border.
        //     When the item is serialized out as xml, its value is "handmade1".
        Handmade1 = 103,
        //
        // 摘要:
        //     Handmade Art Border.
        //     When the item is serialized out as xml, its value is "handmade2".
        Handmade2 = 104,
        //
        // 摘要:
        //     Heart-Shaped Balloon Art Border.
        //     When the item is serialized out as xml, its value is "heartBalloon".
        HeartBalloon = 105,
        //
        // 摘要:
        //     Gray Heart Art Border.
        //     When the item is serialized out as xml, its value is "heartGray".
        HeartGray = 106,
        //
        // 摘要:
        //     Hearts Art Border.
        //     When the item is serialized out as xml, its value is "hearts".
        Hearts = 107,
        //
        // 摘要:
        //     Pattern Art Border.
        //     When the item is serialized out as xml, its value is "heebieJeebies".
        HeebieJeebies = 108,
        //
        // 摘要:
        //     Holly Art Border.
        //     When the item is serialized out as xml, its value is "holly".
        Holly = 109,
        //
        // 摘要:
        //     House Art Border.
        //     When the item is serialized out as xml, its value is "houseFunky".
        HouseFunky = 110,
        //
        // 摘要:
        //     Circular Art Border.
        //     When the item is serialized out as xml, its value is "hypnotic".
        Hypnotic = 111,
        //
        // 摘要:
        //     Ice Cream Cone Art Border.
        //     When the item is serialized out as xml, its value is "iceCreamCones".
        IceCreamCones = 112,
        //
        // 摘要:
        //     Light Bulb Art Border.
        //     When the item is serialized out as xml, its value is "lightBulb".
        LightBulb = 113,
        //
        // 摘要:
        //     Lightning Art Border.
        //     When the item is serialized out as xml, its value is "lightning1".
        Lightning1 = 114,
        //
        // 摘要:
        //     Lightning Art Border.
        //     When the item is serialized out as xml, its value is "lightning2".
        Lightning2 = 115,
        //
        // 摘要:
        //     Map Pins Art Border.
        //     When the item is serialized out as xml, its value is "mapPins".
        MapPins = 116,
        //
        // 摘要:
        //     Maple Leaf Art Border.
        //     When the item is serialized out as xml, its value is "mapleLeaf".
        MapleLeaf = 117,
        //
        // 摘要:
        //     Muffin Art Border.
        //     When the item is serialized out as xml, its value is "mapleMuffins".
        MapleMuffins = 118,
        //
        // 摘要:
        //     Marquee Art Border.
        //     When the item is serialized out as xml, its value is "marquee".
        Marquee = 119,
        //
        // 摘要:
        //     Marquee Art Border.
        //     When the item is serialized out as xml, its value is "marqueeToothed".
        MarqueeToothed = 120,
        //
        // 摘要:
        //     Moon Art Border.
        //     When the item is serialized out as xml, its value is "moons".
        Moons = 121,
        //
        // 摘要:
        //     Mosaic Art Border.
        //     When the item is serialized out as xml, its value is "mosaic".
        Mosaic = 122,
        //
        // 摘要:
        //     Musical Note Art Border.
        //     When the item is serialized out as xml, its value is "musicNotes".
        MusicNotes = 123,
        //
        // 摘要:
        //     Patterned Art Border.
        //     When the item is serialized out as xml, its value is "northwest".
        Northwest = 124,
        //
        // 摘要:
        //     Oval Art Border.
        //     When the item is serialized out as xml, its value is "ovals".
        Ovals = 125,
        //
        // 摘要:
        //     Package Art Border.
        //     When the item is serialized out as xml, its value is "packages".
        Packages = 126,
        //
        // 摘要:
        //     Black Palm Tree Art Border.
        //     When the item is serialized out as xml, its value is "palmsBlack".
        PalmsBlack = 127,
        //
        // 摘要:
        //     Color Palm Tree Art Border.
        //     When the item is serialized out as xml, its value is "palmsColor".
        PalmsColor = 128,
        //
        // 摘要:
        //     Paper Clip Art Border.
        //     When the item is serialized out as xml, its value is "paperClips".
        PaperClips = 129,
        //
        // 摘要:
        //     Papyrus Art Border.
        //     When the item is serialized out as xml, its value is "papyrus".
        Papyrus = 130,
        //
        // 摘要:
        //     Party Favor Art Border.
        //     When the item is serialized out as xml, its value is "partyFavor".
        PartyFavor = 131,
        //
        // 摘要:
        //     Party Glass Art Border.
        //     When the item is serialized out as xml, its value is "partyGlass".
        PartyGlass = 132,
        //
        // 摘要:
        //     Pencils Art Border.
        //     When the item is serialized out as xml, its value is "pencils".
        Pencils = 133,
        //
        // 摘要:
        //     Character Art Border.
        //     When the item is serialized out as xml, its value is "people".
        People = 134,
        //
        // 摘要:
        //     Waving Character Border.
        //     When the item is serialized out as xml, its value is "peopleWaving".
        PeopleWaving = 135,
        //
        // 摘要:
        //     Character With Hat Art Border.
        //     When the item is serialized out as xml, its value is "peopleHats".
        PeopleHats = 136,
        //
        // 摘要:
        //     Poinsettia Art Border.
        //     When the item is serialized out as xml, its value is "poinsettias".
        Poinsettias = 137,
        //
        // 摘要:
        //     Postage Stamp Art Border.
        //     When the item is serialized out as xml, its value is "postageStamp".
        PostageStamp = 138,
        //
        // 摘要:
        //     Pumpkin Art Border.
        //     When the item is serialized out as xml, its value is "pumpkin1".
        Pumpkin1 = 139,
        //
        // 摘要:
        //     Push Pin Art Border.
        //     When the item is serialized out as xml, its value is "pushPinNote2".
        PushPinNote2 = 140,
        //
        // 摘要:
        //     Push Pin Art Border.
        //     When the item is serialized out as xml, its value is "pushPinNote1".
        PushPinNote1 = 141,
        //
        // 摘要:
        //     Pyramid Art Border.
        //     When the item is serialized out as xml, its value is "pyramids".
        Pyramids = 142,
        //
        // 摘要:
        //     Pyramid Art Border.
        //     When the item is serialized out as xml, its value is "pyramidsAbove".
        PyramidsAbove = 143,
        //
        // 摘要:
        //     Quadrants Art Border.
        //     When the item is serialized out as xml, its value is "quadrants".
        Quadrants = 144,
        //
        // 摘要:
        //     Rings Art Border.
        //     When the item is serialized out as xml, its value is "rings".
        Rings = 145,
        //
        // 摘要:
        //     Safari Art Border.
        //     When the item is serialized out as xml, its value is "safari".
        Safari = 146,
        //
        // 摘要:
        //     Saw tooth Art Border.
        //     When the item is serialized out as xml, its value is "sawtooth".
        Sawtooth = 147,
        //
        // 摘要:
        //     Gray Saw tooth Art Border.
        //     When the item is serialized out as xml, its value is "sawtoothGray".
        SawtoothGray = 148,
        //
        // 摘要:
        //     Scared Cat Art Border.
        //     When the item is serialized out as xml, its value is "scaredCat".
        ScaredCat = 149,
        //
        // 摘要:
        //     Umbrella Art Border.
        //     When the item is serialized out as xml, its value is "seattle".
        Seattle = 150,
        //
        // 摘要:
        //     Shadowed Squares Art Border.
        //     When the item is serialized out as xml, its value is "shadowedSquares".
        ShadowedSquares = 151,
        //
        // 摘要:
        //     Shark Tooth Art Border.
        //     When the item is serialized out as xml, its value is "sharksTeeth".
        SharksTeeth = 152,
        //
        // 摘要:
        //     Bird Tracks Art Border.
        //     When the item is serialized out as xml, its value is "shorebirdTracks".
        ShorebirdTracks = 153,
        //
        // 摘要:
        //     Rocket Art Border.
        //     When the item is serialized out as xml, its value is "skyrocket".
        Skyrocket = 154,
        //
        // 摘要:
        //     Snowflake Art Border.
        //     When the item is serialized out as xml, its value is "snowflakeFancy".
        SnowflakeFancy = 155,
        //
        // 摘要:
        //     Snowflake Art Border.
        //     When the item is serialized out as xml, its value is "snowflakes".
        Snowflakes = 156,
        //
        // 摘要:
        //     Sombrero Art Border.
        //     When the item is serialized out as xml, its value is "sombrero".
        Sombrero = 157,
        //
        // 摘要:
        //     Southwest-themed Art Border.
        //     When the item is serialized out as xml, its value is "southwest".
        Southwest = 158,
        //
        // 摘要:
        //     Stars Art Border.
        //     When the item is serialized out as xml, its value is "stars".
        Stars = 159,
        //
        // 摘要:
        //     Stars On Top Art Border.
        //     When the item is serialized out as xml, its value is "starsTop".
        StarsTop = 160,
        //
        // 摘要:
        //     3-D Stars Art Border.
        //     When the item is serialized out as xml, its value is "stars3d".
        Stars3d = 161,
        //
        // 摘要:
        //     Stars Art Border.
        //     When the item is serialized out as xml, its value is "starsBlack".
        StarsBlack = 162,
        //
        // 摘要:
        //     Stars With Shadows Art Border.
        //     When the item is serialized out as xml, its value is "starsShadowed".
        StarsShadowed = 163,
        //
        // 摘要:
        //     Sun Art Border.
        //     When the item is serialized out as xml, its value is "sun".
        Sun = 164,
        //
        // 摘要:
        //     Whirligig Art Border.
        //     When the item is serialized out as xml, its value is "swirligig".
        Swirligig = 165,
        //
        // 摘要:
        //     Torn Paper Art Border.
        //     When the item is serialized out as xml, its value is "tornPaper".
        TornPaper = 166,
        //
        // 摘要:
        //     Black Torn Paper Art Border.
        //     When the item is serialized out as xml, its value is "tornPaperBlack".
        TornPaperBlack = 167,
        //
        // 摘要:
        //     Tree Art Border.
        //     When the item is serialized out as xml, its value is "trees".
        Trees = 168,
        //
        // 摘要:
        //     Triangle Art Border.
        //     When the item is serialized out as xml, its value is "triangleParty".
        TriangleParty = 169,
        //
        // 摘要:
        //     Triangles Art Border.
        //     When the item is serialized out as xml, its value is "triangles".
        Triangles = 170,
        //
        // 摘要:
        //     Tribal Art Border One.
        //     When the item is serialized out as xml, its value is "tribal1".
        Tribal1 = 171,
        //
        // 摘要:
        //     Tribal Art Border Two.
        //     When the item is serialized out as xml, its value is "tribal2".
        Tribal2 = 172,
        //
        // 摘要:
        //     Tribal Art Border Three.
        //     When the item is serialized out as xml, its value is "tribal3".
        Tribal3 = 173,
        //
        // 摘要:
        //     Tribal Art Border Four.
        //     When the item is serialized out as xml, its value is "tribal4".
        Tribal4 = 174,
        //
        // 摘要:
        //     Tribal Art Border Five.
        //     When the item is serialized out as xml, its value is "tribal5".
        Tribal5 = 175,
        //
        // 摘要:
        //     Tribal Art Border Six.
        //     When the item is serialized out as xml, its value is "tribal6".
        Tribal6 = 176,
        //
        // 摘要:
        //     triangle1.
        //     When the item is serialized out as xml, its value is "triangle1".
        Triangle1 = 177,
        //
        // 摘要:
        //     triangle2.
        //     When the item is serialized out as xml, its value is "triangle2".
        Triangle2 = 178,
        //
        // 摘要:
        //     triangleCircle1.
        //     When the item is serialized out as xml, its value is "triangleCircle1".
        TriangleCircle1 = 179,
        //
        // 摘要:
        //     triangleCircle2.
        //     When the item is serialized out as xml, its value is "triangleCircle2".
        TriangleCircle2 = 180,
        //
        // 摘要:
        //     shapes1.
        //     When the item is serialized out as xml, its value is "shapes1".
        Shapes1 = 181,
        //
        // 摘要:
        //     shapes2.
        //     When the item is serialized out as xml, its value is "shapes2".
        Shapes2 = 182,
        //
        // 摘要:
        //     Twisted Lines Art Border.
        //     When the item is serialized out as xml, its value is "twistedLines1".
        TwistedLines1 = 183,
        //
        // 摘要:
        //     Twisted Lines Art Border.
        //     When the item is serialized out as xml, its value is "twistedLines2".
        TwistedLines2 = 184,
        //
        // 摘要:
        //     Vine Art Border.
        //     When the item is serialized out as xml, its value is "vine".
        Vine = 185,
        //
        // 摘要:
        //     Wavy Line Art Border.
        //     When the item is serialized out as xml, its value is "waveline".
        Waveline = 186,
        //
        // 摘要:
        //     Weaving Angles Art Border.
        //     When the item is serialized out as xml, its value is "weavingAngles".
        WeavingAngles = 187,
        //
        // 摘要:
        //     Weaving Braid Art Border.
        //     When the item is serialized out as xml, its value is "weavingBraid".
        WeavingBraid = 188,
        //
        // 摘要:
        //     Weaving Ribbon Art Border.
        //     When the item is serialized out as xml, its value is "weavingRibbon".
        WeavingRibbon = 189,
        //
        // 摘要:
        //     Weaving Strips Art Border.
        //     When the item is serialized out as xml, its value is "weavingStrips".
        WeavingStrips = 190,
        //
        // 摘要:
        //     White Flowers Art Border.
        //     When the item is serialized out as xml, its value is "whiteFlowers".
        WhiteFlowers = 191,
        //
        // 摘要:
        //     Woodwork Art Border.
        //     When the item is serialized out as xml, its value is "woodwork".
        Woodwork = 192,
        //
        // 摘要:
        //     Crisscross Art Border.
        //     When the item is serialized out as xml, its value is "xIllusions".
        XIllusions = 193,
        //
        // 摘要:
        //     Triangle Art Border.
        //     When the item is serialized out as xml, its value is "zanyTriangles".
        ZanyTriangles = 194,
        //
        // 摘要:
        //     Zigzag Art Border.
        //     When the item is serialized out as xml, its value is "zigZag".
        ZigZag = 195,
        //
        // 摘要:
        //     Zigzag stitch.
        //     When the item is serialized out as xml, its value is "zigZagStitch".
        ZigZagStitch = 196
    }

    /// <summary>
    /// Defiines the style of custom tab stop.
    /// </summary>
    public enum TabStopStyle
    {
        /// <summary>
        /// No Tab Stop.
        /// </summary>
        Clear = 0,
        /// <summary>
        /// Left Tab.
        /// </summary>
        Left = 1,
        /// <summary>
        /// Start.
        /// </summary>
        Start = 2,
        /// <summary>
        /// Centered Tab.
        /// </summary>
        Center = 3,
        /// <summary>
        /// Right Tab.
        /// </summary>
        Right = 4,
        /// <summary>
        /// end.
        /// </summary>
        End = 5,
        /// <summary>
        /// Decimal Tab.
        /// </summary>
        Decimal = 6,
        /// <summary>
        /// Bar Tab.
        /// </summary>
        Bar = 7,
        /// <summary>
        /// List Tab.
        /// </summary>
        Number = 8
    }

    /// <summary>
    /// Defines Tab Leader Character enumeration.
    /// </summary>
    public enum TabStopLeader
    {
        /// <summary>
        /// No tab stop leader.
        /// </summary>
        None = 0,
        /// <summary>
        /// Dotted leader line.
        /// </summary>
        Dot = 1,
        /// <summary>
        /// Dashed tab stop leader line.
        /// </summary>
        Hyphen = 2,
        /// <summary>
        /// Solid leader line.
        /// </summary>
        Underscore = 3,
        /// <summary>
        /// Heavy solid leader line.
        /// </summary>
        Heavy = 4,
        /// <summary>
        /// Middle dot leader line.
        /// </summary>
        MiddleDot = 5
    }

    public enum ListNumberStyle
    {
        //
        // 摘要:
        //     Decimal Numbers.
        //     When the item is serialized out as xml, its value is "decimal".
        Decimal = 0,
        //
        // 摘要:
        //     Uppercase Roman Numerals.
        //     When the item is serialized out as xml, its value is "upperRoman".
        UpperRoman = 1,
        //
        // 摘要:
        //     Lowercase Roman Numerals.
        //     When the item is serialized out as xml, its value is "lowerRoman".
        LowerRoman = 2,
        //
        // 摘要:
        //     Uppercase Latin Alphabet.
        //     When the item is serialized out as xml, its value is "upperLetter".
        UpperLetter = 3,
        //
        // 摘要:
        //     Lowercase Latin Alphabet.
        //     When the item is serialized out as xml, its value is "lowerLetter".
        LowerLetter = 4,
        //
        // 摘要:
        //     Ordinal.
        //     When the item is serialized out as xml, its value is "ordinal".
        Ordinal = 5,
        //
        // 摘要:
        //     Cardinal Text.
        //     When the item is serialized out as xml, its value is "cardinalText".
        CardinalText = 6,
        //
        // 摘要:
        //     Ordinal Text.
        //     When the item is serialized out as xml, its value is "ordinalText".
        OrdinalText = 7,
        //
        // 摘要:
        //     Hexadecimal Numbering.
        //     When the item is serialized out as xml, its value is "hex".
        Hex = 8,
        //
        // 摘要:
        //     Chicago Manual of Style.
        //     When the item is serialized out as xml, its value is "chicago".
        Chicago = 9,
        //
        // 摘要:
        //     Ideographs.
        //     When the item is serialized out as xml, its value is "ideographDigital".
        IdeographDigital = 10,
        //
        // 摘要:
        //     Japanese Counting System.
        //     When the item is serialized out as xml, its value is "japaneseCounting".
        JapaneseCounting = 11,
        //
        // 摘要:
        //     AIUEO Order Hiragana.
        //     When the item is serialized out as xml, its value is "aiueo".
        Aiueo = 12,
        //
        // 摘要:
        //     Iroha Ordered Katakana.
        //     When the item is serialized out as xml, its value is "iroha".
        Iroha = 13,
        //
        // 摘要:
        //     Double Byte Arabic Numerals.
        //     When the item is serialized out as xml, its value is "decimalFullWidth".
        DecimalFullWidth = 14,
        //
        // 摘要:
        //     Single Byte Arabic Numerals.
        //     When the item is serialized out as xml, its value is "decimalHalfWidth".
        DecimalHalfWidth = 15,
        //
        // 摘要:
        //     Japanese Legal Numbering.
        //     When the item is serialized out as xml, its value is "japaneseLegal".
        JapaneseLegal = 16,
        //
        // 摘要:
        //     Japanese Digital Ten Thousand Counting System.
        //     When the item is serialized out as xml, its value is "japaneseDigitalTenThousand".
        JapaneseDigitalTenThousand = 17,
        //
        // 摘要:
        //     Decimal Numbers Enclosed in a Circle.
        //     When the item is serialized out as xml, its value is "decimalEnclosedCircle".
        DecimalEnclosedCircle = 18,
        //
        // 摘要:
        //     Double Byte Arabic Numerals Alternate.
        //     When the item is serialized out as xml, its value is "decimalFullWidth2".
        DecimalFullWidth2 = 19,
        //
        // 摘要:
        //     Full-Width AIUEO Order Hiragana.
        //     When the item is serialized out as xml, its value is "aiueoFullWidth".
        AiueoFullWidth = 20,
        //
        // 摘要:
        //     Full-Width Iroha Ordered Katakana.
        //     When the item is serialized out as xml, its value is "irohaFullWidth".
        IrohaFullWidth = 21,
        //
        // 摘要:
        //     Initial Zero Arabic Numerals.
        //     When the item is serialized out as xml, its value is "decimalZero".
        DecimalZero = 22,
        //
        // 摘要:
        //     Bullet.
        //     When the item is serialized out as xml, its value is "bullet".
        Bullet = 23,
        //
        // 摘要:
        //     Korean Ganada Numbering.
        //     When the item is serialized out as xml, its value is "ganada".
        Ganada = 24,
        //
        // 摘要:
        //     Korean Chosung Numbering.
        //     When the item is serialized out as xml, its value is "chosung".
        Chosung = 25,
        //
        // 摘要:
        //     Decimal Numbers Followed by a Period.
        //     When the item is serialized out as xml, its value is "decimalEnclosedFullstop".
        DecimalEnclosedFullstop = 26,
        //
        // 摘要:
        //     Decimal Numbers Enclosed in Parenthesis.
        //     When the item is serialized out as xml, its value is "decimalEnclosedParen".
        DecimalEnclosedParen = 27,
        //
        // 摘要:
        //     Decimal Numbers Enclosed in a Circle.
        //     When the item is serialized out as xml, its value is "decimalEnclosedCircleChinese".
        DecimalEnclosedCircleChinese = 28,
        //
        // 摘要:
        //     Ideographs Enclosed in a Circle.
        //     When the item is serialized out as xml, its value is "ideographEnclosedCircle".
        IdeographEnclosedCircle = 29,
        //
        // 摘要:
        //     Traditional Ideograph Format.
        //     When the item is serialized out as xml, its value is "ideographTraditional".
        IdeographTraditional = 30,
        //
        // 摘要:
        //     Zodiac Ideograph Format.
        //     When the item is serialized out as xml, its value is "ideographZodiac".
        IdeographZodiac = 31,
        //
        // 摘要:
        //     Traditional Zodiac Ideograph Format.
        //     When the item is serialized out as xml, its value is "ideographZodiacTraditional".
        IdeographZodiacTraditional = 32,
        //
        // 摘要:
        //     Taiwanese Counting System.
        //     When the item is serialized out as xml, its value is "taiwaneseCounting".
        TaiwaneseCounting = 33,
        //
        // 摘要:
        //     Traditional Legal Ideograph Format.
        //     When the item is serialized out as xml, its value is "ideographLegalTraditional".
        IdeographLegalTraditional = 34,
        //
        // 摘要:
        //     Taiwanese Counting Thousand System.
        //     When the item is serialized out as xml, its value is "taiwaneseCountingThousand".
        TaiwaneseCountingThousand = 35,
        //
        // 摘要:
        //     Taiwanese Digital Counting System.
        //     When the item is serialized out as xml, its value is "taiwaneseDigital".
        TaiwaneseDigital = 36,
        //
        // 摘要:
        //     Chinese Counting System.
        //     When the item is serialized out as xml, its value is "chineseCounting".
        ChineseCounting = 37,
        //
        // 摘要:
        //     Chinese Legal Simplified Format.
        //     When the item is serialized out as xml, its value is "chineseLegalSimplified".
        ChineseLegalSimplified = 38,
        //
        // 摘要:
        //     Chinese Counting Thousand System.
        //     When the item is serialized out as xml, its value is "chineseCountingThousand".
        ChineseCountingThousand = 39,
        //
        // 摘要:
        //     Korean Digital Counting System.
        //     When the item is serialized out as xml, its value is "koreanDigital".
        KoreanDigital = 40,
        //
        // 摘要:
        //     Korean Counting System.
        //     When the item is serialized out as xml, its value is "koreanCounting".
        KoreanCounting = 41,
        //
        // 摘要:
        //     Korean Legal Numbering.
        //     When the item is serialized out as xml, its value is "koreanLegal".
        KoreanLegal = 42,
        //
        // 摘要:
        //     Korean Digital Counting System Alternate.
        //     When the item is serialized out as xml, its value is "koreanDigital2".
        KoreanDigital2 = 43,
        //
        // 摘要:
        //     Vietnamese Numerals.
        //     When the item is serialized out as xml, its value is "vietnameseCounting".
        VietnameseCounting = 44,
        //
        // 摘要:
        //     Lowercase Russian Alphabet.
        //     When the item is serialized out as xml, its value is "russianLower".
        RussianLower = 45,
        //
        // 摘要:
        //     Uppercase Russian Alphabet.
        //     When the item is serialized out as xml, its value is "russianUpper".
        RussianUpper = 46,
        //
        // 摘要:
        //     No Numbering.
        //     When the item is serialized out as xml, its value is "none".
        None = 47,
        //
        // 摘要:
        //     Number With Dashes.
        //     When the item is serialized out as xml, its value is "numberInDash".
        NumberInDash = 48,
        //
        // 摘要:
        //     Hebrew Numerals.
        //     When the item is serialized out as xml, its value is "hebrew1".
        Hebrew1 = 49,
        //
        // 摘要:
        //     Hebrew Alphabet.
        //     When the item is serialized out as xml, its value is "hebrew2".
        Hebrew2 = 50,
        //
        // 摘要:
        //     Arabic Alphabet.
        //     When the item is serialized out as xml, its value is "arabicAlpha".
        ArabicAlpha = 51,
        //
        // 摘要:
        //     Arabic Abjad Numerals.
        //     When the item is serialized out as xml, its value is "arabicAbjad".
        ArabicAbjad = 52,
        //
        // 摘要:
        //     Hindi Vowels.
        //     When the item is serialized out as xml, its value is "hindiVowels".
        HindiVowels = 53,
        //
        // 摘要:
        //     Hindi Consonants.
        //     When the item is serialized out as xml, its value is "hindiConsonants".
        HindiConsonants = 54,
        //
        // 摘要:
        //     Hindi Numbers.
        //     When the item is serialized out as xml, its value is "hindiNumbers".
        HindiNumbers = 55,
        //
        // 摘要:
        //     Hindi Counting System.
        //     When the item is serialized out as xml, its value is "hindiCounting".
        HindiCounting = 56,
        //
        // 摘要:
        //     Thai Letters.
        //     When the item is serialized out as xml, its value is "thaiLetters".
        ThaiLetters = 57,
        //
        // 摘要:
        //     Thai Numerals.
        //     When the item is serialized out as xml, its value is "thaiNumbers".
        ThaiNumbers = 58,
        //
        // 摘要:
        //     Thai Counting System.
        //     When the item is serialized out as xml, its value is "thaiCounting".
        ThaiCounting = 59,
        //
        // 摘要:
        //     bahtText.
        //     When the item is serialized out as xml, its value is "bahtText".
        //     This item is only available in Office 2010 and later.
        BahtText = 60,
        //
        // 摘要:
        //     dollarText.
        //     When the item is serialized out as xml, its value is "dollarText".
        //     This item is only available in Office 2010 and later.
        DollarText = 61,
        //
        // 摘要:
        //     custom.
        //     When the item is serialized out as xml, its value is "custom".
        //     This item is only available in Office 2010 and later.
        Custom = 62
    }
}
