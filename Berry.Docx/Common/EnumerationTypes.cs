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
        Break,
        Tab,
        DeletedRange,
        InsertedRange,
        DeletedTextRange
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
        /// <summary>
        /// Double Line Border.
        /// </summary>
        Double = 4,
        /// <summary>
        /// Dotted Line Border.
        /// </summary>
        Dotted = 5,
        /// <summary>
        /// Dashed Line Border.
        /// </summary>
        Dashed = 6,
        /// <summary>
        /// Dot Dash Line Border.
        /// </summary>
        DotDash = 7,
        /// <summary>
        /// Dot Dot Dash Line Border.
        /// </summary>
        DotDotDash = 8,
        /// <summary>
        /// Triple Line Border.
        /// </summary>
        Triple = 9,
        /// <summary>
        /// Thin, Thick Line Border.
        /// </summary>
        ThinThickSmallGap = 10,
        /// <summary>
        /// Thick, Thin Line Border.
        /// </summary>
        ThickThinSmallGap = 11,
        /// <summary>
        /// Thin, Thick, Thin Line Border.
        /// </summary>
        ThinThickThinSmallGap = 12,
        /// <summary>
        /// Thin, Thick Line Border.
        /// </summary>
        ThinThickMediumGap = 13,
        /// <summary>
        /// Thick, Thin Line Border.
        /// </summary>
        ThickThinMediumGap = 14,
        /// <summary>
        /// Thin, Thick, Thin Line Border.
        /// </summary>
        ThinThickThinMediumGap = 15,
        /// <summary>
        /// Thin, Thick Line Border.
        /// </summary>
        ThinThickLargeGap = 16,
        /// <summary>
        /// Thick, Thin Line Border.
        /// </summary>
        ThickThinLargeGap = 17,
        /// <summary>
        /// Thin, Thick, Thin Line Border.
        /// </summary>
        ThinThickThinLargeGap = 18,
        /// <summary>
        /// Wavy Line Border.
        /// </summary>
        Wave = 19,
        /// <summary>
        /// Double Wave Line Border.
        /// </summary>
        DoubleWave = 20,
        /// <summary>
        /// Dashed Line Border.
        /// </summary>
        DashSmallGap = 21,
        /// <summary>
        /// Dash Dot Strokes Line Border.
        /// </summary>
        DashDotStroked = 22,
        /// <summary>
        /// 3D Embossed Line Border.
        /// </summary>
        ThreeDEmboss = 23,
        /// <summary>
        /// 3D Engraved Line Border.
        /// </summary>
        ThreeDEngrave = 24,
        /// <summary>
        /// Outset Line Border.
        /// </summary>
        Outset = 25,
        /// <summary>
        /// Inset Line Border.
        /// </summary>
        Inset = 26,
        /// <summary>
        /// Apples Art Border.
        /// </summary>
        Apples = 27,
        /// <summary>
        /// Arched Scallops Art Border.
        /// </summary>
        ArchedScallops = 28,
        /// <summary>
        /// Baby Pacifier Art Border.
        /// </summary>
        BabyPacifier = 29,
        /// <summary>
        /// Baby Rattle Art Border.
        /// </summary>
        BabyRattle = 30,
        /// <summary>
        /// Three Color Balloons Art Border.
        /// </summary>
        Balloons3Colors = 31,
        /// <summary>
        /// Hot Air Balloons Art Border.
        /// </summary>
        BalloonsHotAir = 32,
        /// <summary>
        /// Black Dash Art Border.
        /// </summary>
        BasicBlackDashes = 33,
        /// <summary>
        /// Black Dot Art Border.
        /// </summary>
        BasicBlackDots = 34,
        /// <summary>
        /// Black Square Art Border.
        /// </summary>
        BasicBlackSquares = 35,
        /// <summary>
        /// Thin Line Art Border.
        /// </summary>
        BasicThinLines = 36,
        /// <summary>
        /// White Dash Art Border.
        /// </summary>
        BasicWhiteDashes = 37,
        /// <summary>
        /// White Dot Art Border.
        /// </summary>
        BasicWhiteDots = 38,
        /// <summary>
        /// White Square Art Border.
        /// </summary>
        BasicWhiteSquares = 39,
        /// <summary>
        /// Wide Inline Art Border.
        /// </summary>
        BasicWideInline = 40,
        /// <summary>
        /// Wide Midline Art Border.
        /// </summary>
        BasicWideMidline = 41,
        /// <summary>
        /// Wide Outline Art Border.
        /// </summary>
        BasicWideOutline = 42,
        /// <summary>
        /// Bats Art Border.
        /// </summary>
        Bats = 43,
        /// <summary>
        /// Birds Art Border.
        /// </summary>
        Birds = 44,
        /// <summary>
        /// Birds Flying Art Border.
        /// </summary>
        BirdsFlight = 45,
        /// <summary>
        /// Cabin Art Border.
        /// </summary>
        Cabins = 46,
        /// <summary>
        /// Cake Art Border.
        /// </summary>
        CakeSlice = 47,
        /// <summary>
        /// Candy Corn Art Border.
        /// </summary>
        CandyCorn = 48,
        /// <summary>
        /// Knot Work Art Border.
        /// </summary>
        CelticKnotwork = 49,
        /// <summary>
        /// Certificate Banner Art Border.
        /// </summary>
        CertificateBanner = 50,
        /// <summary>
        /// Chain Link Art Border.
        /// </summary>
        ChainLink = 51,
        /// <summary>
        /// Champagne Bottle Art Border.
        /// </summary>
        ChampagneBottle = 52,
        /// <summary>
        /// Black and White Bar Art Border.
        /// </summary>
        CheckedBarBlack = 53,
        /// <summary>
        /// Color Checked Bar Art Border.
        /// </summary>
        CheckedBarColor = 54,
        /// <summary>
        /// Checkerboard Art Border.
        /// </summary>
        Checkered = 55,
        /// <summary>
        /// Christmas Tree Art Border.
        /// </summary>
        ChristmasTree = 56,
        /// <summary>
        /// Circles And Lines Art Border.
        /// </summary>
        CirclesLines = 57,
        /// <summary>
        /// Circles and Rectangles Art Border.
        /// </summary>
        CirclesRectangles = 58,
        /// <summary>
        /// Wave Art Border.
        /// </summary>
        ClassicalWave = 59,
        /// <summary>
        /// Clocks Art Border.
        /// </summary>
        Clocks = 60,
        /// <summary>
        /// Compass Art Border.
        /// </summary>
        Compass = 61,
        /// <summary>
        /// Confetti Art Border.
        /// </summary>
        Confetti = 62,
        /// <summary>
        /// Confetti Art Border.
        /// </summary>
        ConfettiGrays = 63,
        /// <summary>
        /// Confetti Art Border.
        /// </summary>
        ConfettiOutline = 64,
        /// <summary>
        /// Confetti Streamers Art Border.
        /// </summary>
        ConfettiStreamers = 65,
        /// <summary>
        /// Confetti Art Border.
        /// </summary>
        ConfettiWhite = 66,
        /// <summary>
        /// Corner Triangle Art Border.
        /// </summary>
        CornerTriangles = 67,
        /// <summary>
        /// Dashed Line Art Border.
        /// </summary>
        CouponCutoutDashes = 68,
        /// <summary>
        /// Dotted Line Art Border.
        /// </summary>
        CouponCutoutDots = 69,
        /// <summary>
        /// Maze Art Border.
        /// </summary>
        CrazyMaze = 70,
        /// <summary>
        /// Butterfly Art Border.
        /// </summary>
        CreaturesButterfly = 71,
        /// <summary>
        /// Fish Art Border.
        /// </summary>
        CreaturesFish = 72,
        /// <summary>
        /// Insects Art Border.
        /// </summary>
        CreaturesInsects = 73,
        /// <summary>
        /// Ladybug Art Border.
        /// </summary>
        CreaturesLadyBug = 74,
        /// <summary>
        /// Cross-stitch Art Border.
        /// </summary>
        CrossStitch = 75,
        /// <summary>
        /// Cupid Art Border.
        /// </summary>
        Cup = 76,
        /// <summary>
        /// Archway Art Border.
        /// </summary>
        DecoArch = 77,
        /// <summary>
        /// Color Archway Art Border.
        /// </summary>
        DecoArchColor = 78,
        /// <summary>
        /// Blocks Art Border.
        /// </summary>
        DecoBlocks = 79,
        /// <summary>
        /// Gray Diamond Art Border.
        /// </summary>
        DiamondsGray = 80,
        /// <summary>
        /// Double D Art Border.
        /// </summary>
        DoubleD = 81,
        /// <summary>
        /// Diamond Art Border.
        /// </summary>
        DoubleDiamonds = 82,
        /// <summary>
        /// Earth Art Border.
        /// </summary>
        Earth1 = 83,
        /// <summary>
        /// Earth Art Border.
        /// </summary>
        Earth2 = 84,
        /// <summary>
        /// Shadowed Square Art Border.
        /// </summary>
        EclipsingSquares1 = 85,
        /// <summary>
        /// Shadowed Square Art Border.
        /// </summary>
        EclipsingSquares2 = 86,
        /// <summary>
        /// Painted Egg Art Border.
        /// </summary>
        EggsBlack = 87,
        /// <summary>
        /// Fans Art Border.
        /// </summary>
        Fans = 88,
        /// <summary>
        /// Film Reel Art Border.
        /// </summary>
        Film = 89,
        /// <summary>
        /// Firecracker Art Border.
        /// </summary>
        Firecrackers = 90,
        /// <summary>
        /// Flowers Art Border.
        /// </summary>
        FlowersBlockPrint = 91,
        /// <summary>
        /// Daisy Art Border.
        /// </summary>
        FlowersDaisies = 92,
        /// <summary>
        /// Flowers Art Border.
        /// </summary>
        FlowersModern1 = 93,
        /// <summary>
        /// Flowers Art Border.
        /// </summary>
        FlowersModern2 = 94,
        /// <summary>
        /// Pansy Art Border.
        /// </summary>
        FlowersPansy = 95,
        /// <summary>
        /// Red Rose Art Border.
        /// </summary>
        FlowersRedRose = 96,
        /// <summary>
        /// Roses Art Border.
        /// </summary>
        FlowersRoses = 97,
        /// <summary>
        /// Flowers in a Teacup Art Border.
        /// </summary>
        FlowersTeacup = 98,
        /// <summary>
        /// Small Flower Art Border.
        /// </summary>
        FlowersTiny = 99,
        /// <summary>
        /// Gems Art Border.
        /// </summary>
        Gems = 100,
        /// <summary>
        /// Gingerbread Man Art Border.
        /// </summary>
        GingerbreadMan = 101,
        /// <summary>
        /// Triangle Gradient Art Border.
        /// </summary>
        Gradient = 102,
        /// <summary>
        /// Handmade Art Border.
        /// </summary>
        Handmade1 = 103,
        /// <summary>
        /// Handmade Art Border.
        /// </summary>
        Handmade2 = 104,
        /// <summary>
        /// Heart-Shaped Balloon Art Border.
        /// </summary>
        HeartBalloon = 105,
        /// <summary>
        /// Gray Heart Art Border.
        /// </summary>
        HeartGray = 106,
        /// <summary>
        /// Hearts Art Border.
        /// </summary>
        Hearts = 107,
        /// <summary>
        /// Pattern Art Border.
        /// </summary>
        HeebieJeebies = 108,
        /// <summary>
        /// Holly Art Border.
        /// </summary>
        Holly = 109,
        /// <summary>
        /// House Art Border.
        /// </summary>
        HouseFunky = 110,
        /// <summary>
        /// Circular Art Border.
        /// </summary>
        Hypnotic = 111,
        /// <summary>
        /// Ice Cream Cone Art Border.
        /// </summary>
        IceCreamCones = 112,
        /// <summary>
        /// Light Bulb Art Border.
        /// </summary>
        LightBulb = 113,
        /// <summary>
        /// Lightning Art Border.
        /// </summary>
        Lightning1 = 114,
        /// <summary>
        /// Lightning Art Border.
        /// </summary>
        Lightning2 = 115,
        /// <summary>
        /// Map Pins Art Border.
        /// </summary>
        MapPins = 116,
        /// <summary>
        /// Maple Leaf Art Border.
        /// </summary>
        MapleLeaf = 117,
        /// <summary>
        /// Muffin Art Border.
        /// </summary>
        MapleMuffins = 118,
        /// <summary>
        /// Marquee Art Border.
        /// </summary>
        Marquee = 119,
        /// <summary>
        /// Marquee Art Border.
        /// </summary>
        MarqueeToothed = 120,
        /// <summary>
        /// Moon Art Border.
        /// </summary>
        Moons = 121,
        /// <summary>
        /// Mosaic Art Border.
        /// </summary>
        Mosaic = 122,
        /// <summary>
        /// Musical Note Art Border.
        /// </summary>
        MusicNotes = 123,
        /// <summary>
        /// Patterned Art Border.
        /// </summary>
        Northwest = 124,
        /// <summary>
        /// Oval Art Border.
        /// </summary>
        Ovals = 125,
        /// <summary>
        /// Package Art Border.
        /// </summary>
        Packages = 126,
        /// <summary>
        /// Black Palm Tree Art Border.
        /// </summary>
        PalmsBlack = 127,
        /// <summary>
        /// Color Palm Tree Art Border.
        /// </summary>
        PalmsColor = 128,
        /// <summary>
        /// Paper Clip Art Border.
        /// </summary>
        PaperClips = 129,
        /// <summary>
        /// Papyrus Art Border.
        /// </summary>
        Papyrus = 130,
        /// <summary>
        /// Party Favor Art Border.
        /// </summary>
        PartyFavor = 131,
        /// <summary>
        /// Party Glass Art Border.
        /// </summary>
        PartyGlass = 132,
        /// <summary>
        /// Pencils Art Border.
        /// </summary>
        Pencils = 133,
        /// <summary>
        /// Character Art Border.
        /// </summary>
        People = 134,
        /// <summary>
        /// Waving Character Border.
        /// </summary>
        PeopleWaving = 135,
        /// <summary>
        /// Character With Hat Art Border.
        /// </summary>
        PeopleHats = 136,
        /// <summary>
        /// Poinsettia Art Border.
        /// </summary>
        Poinsettias = 137,
        /// <summary>
        /// Postage Stamp Art Border.
        /// </summary>
        PostageStamp = 138,
        /// <summary>
        /// Pumpkin Art Border.
        /// </summary>
        Pumpkin1 = 139,
        /// <summary>
        /// Push Pin Art Border.
        /// </summary>
        PushPinNote2 = 140,
        /// <summary>
        /// Push Pin Art Border.
        /// </summary>
        PushPinNote1 = 141,
        /// <summary>
        /// Pyramid Art Border.
        /// </summary>
        Pyramids = 142,
        /// <summary>
        /// Pyramid Art Border.
        /// </summary>
        PyramidsAbove = 143,
        /// <summary>
        /// Quadrants Art Border.
        /// </summary>
        Quadrants = 144,
        /// <summary>
        /// Rings Art Border.
        /// </summary>
        Rings = 145,
        /// <summary>
        /// Safari Art Border.
        /// </summary>
        Safari = 146,
        /// <summary>
        /// Saw tooth Art Border.
        /// </summary>
        Sawtooth = 147,
        /// <summary>
        /// Gray Saw tooth Art Border.
        /// </summary>
        SawtoothGray = 148,
        /// <summary>
        /// Scared Cat Art Border.
        /// </summary>
        ScaredCat = 149,
        /// <summary>
        /// Umbrella Art Border.
        /// </summary>
        Seattle = 150,
        /// <summary>
        /// Shadowed Squares Art Border.
        /// </summary>
        ShadowedSquares = 151,
        /// <summary>
        /// Shark Tooth Art Border.
        /// </summary>
        SharksTeeth = 152,
        /// <summary>
        /// Bird Tracks Art Border.
        /// </summary>
        ShorebirdTracks = 153,
        /// <summary>
        /// Rocket Art Border.
        /// </summary>
        Skyrocket = 154,
        /// <summary>
        /// Snowflake Art Border.
        /// </summary>
        SnowflakeFancy = 155,
        /// <summary>
        /// Snowflake Art Border.
        /// </summary>
        Snowflakes = 156,
        /// <summary>
        /// Sombrero Art Border.
        /// </summary>
        Sombrero = 157,
        /// <summary>
        /// Southwest-themed Art Border.
        /// </summary>
        Southwest = 158,
        /// <summary>
        /// Stars Art Border.
        /// </summary>
        Stars = 159,
        /// <summary>
        /// Stars On Top Art Border.
        /// </summary>
        StarsTop = 160,
        /// <summary>
        /// 3-D Stars Art Border.
        /// </summary>
        Stars3d = 161,
        /// <summary>
        /// Stars Art Border.
        /// </summary>
        StarsBlack = 162,
        /// <summary>
        /// Stars With Shadows Art Border.
        /// </summary>
        StarsShadowed = 163,
        /// <summary>
        /// Sun Art Border.
        /// </summary>
        Sun = 164,
        /// <summary>
        /// Whirligig Art Border.
        /// </summary>
        Swirligig = 165,
        /// <summary>
        /// Torn Paper Art Border.
        /// </summary>
        TornPaper = 166,
        /// <summary>
        /// Black Torn Paper Art Border.
        /// </summary>
        TornPaperBlack = 167,
        /// <summary>
        /// Tree Art Border.
        /// </summary>
        Trees = 168,
        /// <summary>
        /// Triangle Art Border.
        /// </summary>
        TriangleParty = 169,
        /// <summary>
        /// Triangles Art Border.
        /// </summary>
        Triangles = 170,
        /// <summary>
        /// Tribal Art Border One.
        /// </summary>
        Tribal1 = 171,
        /// <summary>
        /// Tribal Art Border Two.
        /// </summary>
        Tribal2 = 172,
        /// <summary>
        /// Tribal Art Border Three.
        /// </summary>
        Tribal3 = 173,
        /// <summary>
        /// Tribal Art Border Four.
        /// </summary>
        Tribal4 = 174,
        /// <summary>
        /// Tribal Art Border Five.
        /// </summary>
        Tribal5 = 175,
        /// <summary>
        /// Tribal Art Border Six.
        /// </summary>
        Tribal6 = 176,
        /// <summary>
        /// triangle1.
        /// </summary>
        Triangle1 = 177,
        /// <summary>
        /// triangle2.
        /// </summary>
        Triangle2 = 178,
        /// <summary>
        /// triangleCircle1.
        /// </summary>
        TriangleCircle1 = 179,
        /// <summary>
        /// triangleCircle2.
        /// </summary>
        TriangleCircle2 = 180,
        /// <summary>
        /// shapes1.
        /// </summary>
        Shapes1 = 181,
        /// <summary>
        /// shapes2.
        /// </summary>
        Shapes2 = 182,
        /// <summary>
        /// Twisted Lines Art Border.
        /// </summary>
        TwistedLines1 = 183,
        /// <summary>
        /// Twisted Lines Art Border.
        /// </summary>
        TwistedLines2 = 184,
        /// <summary>
        /// Vine Art Border.
        /// </summary>
        Vine = 185,
        /// <summary>
        /// Wavy Line Art Border.
        /// </summary>
        Waveline = 186,
        /// <summary>
        /// Weaving Angles Art Border.
        /// </summary>
        WeavingAngles = 187,
        /// <summary>
        /// Weaving Braid Art Border.
        /// </summary>
        WeavingBraid = 188,
        /// <summary>
        /// Weaving Ribbon Art Border.
        /// </summary>
        WeavingRibbon = 189,
        /// <summary>
        /// Weaving Strips Art Border.
        /// </summary>
        WeavingStrips = 190,
        /// <summary>
        /// White Flowers Art Border.
        /// </summary>
        WhiteFlowers = 191,
        /// <summary>
        /// Woodwork Art Border.
        /// </summary>
        Woodwork = 192,
        /// <summary>
        /// Crisscross Art Border.
        /// </summary>
        XIllusions = 193,
        /// <summary>
        /// Triangle Art Border.
        /// </summary>
        ZanyTriangles = 194,
        /// <summary>
        /// Zigzag Art Border.
        /// </summary>
        ZigZag = 195,
        /// <summary>
        /// Zigzag stitch.
        /// </summary>
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

    /// <summary>
    /// Defines the ListNumberStyle enumeration.
    /// </summary>
    public enum ListNumberStyle
    {
        /// <summary>
        /// Decimal Numbers.
        /// </summary>
        Decimal = 0,
        /// <summary>
        /// Uppercase Roman Numerals.
        /// </summary>
        UpperRoman = 1,
        /// <summary>
        /// Lowercase Roman Numerals.
        /// </summary>
        LowerRoman = 2,
        /// <summary>
        /// Uppercase Latin Alphabet.
        /// </summary>
        UpperLetter = 3,
        /// <summary>
        /// Lowercase Latin Alphabet.
        /// </summary>
        LowerLetter = 4,
        /// <summary>
        /// Ordinal.
        /// </summary>
        Ordinal = 5,
        /// <summary>
        /// Cardinal Text.
        /// </summary>
        CardinalText = 6,
        /// <summary>
        /// Ordinal Text.
        /// </summary>
        OrdinalText = 7,
        /// <summary>
        /// Hexadecimal Numbering.
        /// </summary>
        Hex = 8,
        /// <summary>
        /// Chicago Manual of Style.
        /// </summary>
        Chicago = 9,
        /// <summary>
        /// Ideographs.
        /// </summary>
        IdeographDigital = 10,
        /// <summary>
        /// Japanese Counting System.
        /// </summary>
        JapaneseCounting = 11,
        /// <summary>
        /// AIUEO Order Hiragana.
        /// </summary>
        Aiueo = 12,
        /// <summary>
        /// Iroha Ordered Katakana.
        /// </summary>
        Iroha = 13,
        /// <summary>
        /// Double Byte Arabic Numerals.
        /// </summary>
        DecimalFullWidth = 14,
        /// <summary>
        /// Single Byte Arabic Numerals.
        /// </summary>
        DecimalHalfWidth = 15,
        /// <summary>
        /// Japanese Legal Numbering.
        /// </summary>
        JapaneseLegal = 16,
        /// <summary>
        /// Japanese Digital Ten Thousand Counting System.
        /// </summary>
        JapaneseDigitalTenThousand = 17,
        /// <summary>
        /// Decimal Numbers Enclosed in a Circle.
        /// </summary>
        DecimalEnclosedCircle = 18,
        /// <summary>
        /// Double Byte Arabic Numerals Alternate.
        /// </summary>
        DecimalFullWidth2 = 19,
        /// <summary>
        /// Full-Width AIUEO Order Hiragana.
        /// </summary>
        AiueoFullWidth = 20,
        /// <summary>
        /// Full-Width Iroha Ordered Katakana.
        /// </summary>
        IrohaFullWidth = 21,
        /// <summary>
        /// Initial Zero Arabic Numerals.
        /// </summary>
        DecimalZero = 22,
        /// <summary>
        /// Bullet.
        /// </summary>
        Bullet = 23,
        /// <summary>
        /// Korean Ganada Numbering.
        /// </summary>
        Ganada = 24,
        /// <summary>
        /// Korean Chosung Numbering.
        /// </summary>
        Chosung = 25,
        /// <summary>
        /// Decimal Numbers Followed by a Period.
        /// </summary>
        DecimalEnclosedFullstop = 26,
        /// <summary>
        /// Decimal Numbers Enclosed in Parenthesis.
        /// </summary>
        DecimalEnclosedParen = 27,
        /// <summary>
        /// Decimal Numbers Enclosed in a Circle.
        /// </summary>
        DecimalEnclosedCircleChinese = 28,
        /// <summary>
        /// Ideographs Enclosed in a Circle.
        /// </summary>
        IdeographEnclosedCircle = 29,
        /// <summary>
        /// Traditional Ideograph Format.
        /// </summary>
        IdeographTraditional = 30,
        /// <summary>
        /// Zodiac Ideograph Format.
        /// </summary>
        IdeographZodiac = 31,
        /// <summary>
        /// Traditional Zodiac Ideograph Format.
        /// </summary>
        IdeographZodiacTraditional = 32,
        /// <summary>
        /// Taiwanese Counting System.
        /// </summary>
        TaiwaneseCounting = 33,
        /// <summary>
        /// Traditional Legal Ideograph Format.
        /// </summary>
        IdeographLegalTraditional = 34,
        /// <summary>
        /// Taiwanese Counting Thousand System.
        /// </summary>
        TaiwaneseCountingThousand = 35,
        /// <summary>
        /// Taiwanese Digital Counting System.
        /// </summary>
        TaiwaneseDigital = 36,
        /// <summary>
        /// Chinese Counting System.
        /// </summary>
        ChineseCounting = 37,
        /// <summary>
        /// Chinese Legal Simplified Format.
        /// </summary>
        ChineseLegalSimplified = 38,
        /// <summary>
        /// Chinese Counting Thousand System.
        /// </summary>
        ChineseCountingThousand = 39,
        /// <summary>
        /// Korean Digital Counting System.
        /// </summary>
        KoreanDigital = 40,
        /// <summary>
        /// Korean Counting System.
        /// </summary>
        KoreanCounting = 41,
        /// <summary>
        /// Korean Legal Numbering.
        /// </summary>
        KoreanLegal = 42,
        /// <summary>
        /// Korean Digital Counting System Alternate.
        /// </summary>
        KoreanDigital2 = 43,
        /// <summary>
        /// Vietnamese Numerals.
        /// </summary>
        VietnameseCounting = 44,
        /// <summary>
        /// Lowercase Russian Alphabet.
        /// </summary>
        RussianLower = 45,
        /// <summary>
        /// Uppercase Russian Alphabet.
        /// </summary>
        RussianUpper = 46,
        /// <summary>
        /// No Numbering.
        /// </summary>
        None = 47,
        /// <summary>
        /// Number With Dashes.
        /// </summary>
        NumberInDash = 48,
        /// <summary>
        /// Hebrew Numerals.
        /// </summary>
        Hebrew1 = 49,
        /// <summary>
        /// Hebrew Alphabet.
        /// </summary>
        Hebrew2 = 50,
        /// <summary>
        /// Arabic Alphabet.
        /// </summary>
        ArabicAlpha = 51,
        /// <summary>
        /// Arabic Abjad Numerals.
        /// </summary>
        ArabicAbjad = 52,
        /// <summary>
        /// Hindi Vowels.
        /// </summary>
        HindiVowels = 53,
        /// <summary>
        /// Hindi Consonants.
        /// </summary>
        HindiConsonants = 54,
        /// <summary>
        /// Hindi Numbers.
        /// </summary>
        HindiNumbers = 55,
        /// <summary>
        /// Hindi Counting System.
        /// </summary>
        HindiCounting = 56,
        /// <summary>
        /// Thai Letters.
        /// </summary>
        ThaiLetters = 57,
        /// <summary>
        /// Thai Numerals.
        /// </summary>
        ThaiNumbers = 58,
        /// <summary>
        /// Thai Counting System.
        /// </summary>
        ThaiCounting = 59,
        /// <summary>
        /// bahtText. This item is only available in Office 2010 and later.
        /// </summary>
        BahtText = 60,
        /// <summary>
        /// dollarText. This item is only available in Office 2010 and later.
        /// </summary>
        DollarText = 61,
        /// <summary>
        /// custom. This item is only available in Office 2010 and later.
        /// </summary>
        Custom = 62
    }

    /// <summary>
    /// Defines the built-in list style enumeration.
    /// </summary>
    public enum BuiltInListStyle
    {
        /// <summary>
        /// <para>1 -------</para>
        /// <para>1.1 -----</para>
        /// <para>1.1.1 ---</para>
        /// </summary>
        Style1 = 1,
        /// <summary>
        /// <para>1. -------</para>
        /// <para>1.1. -----</para>
        /// <para>1.1.1. ---</para>
        /// </summary>
        Style2 = 2,
        /// <summary>
        /// <para>第1章 -------</para>
        /// <para>1.1 -----</para>
        /// <para>1.1.1 ---</para>
        /// </summary>
        Style3 = 3,
        /// <summary>
        /// <para>一 -------</para>
        /// <para>1.1 -----</para>
        /// <para>1.1.1 ---</para>
        /// </summary>
        Style4 = 4
    }

    /// <summary>
    /// Defines the list number alignment enumeration.
    /// </summary>
    public enum ListNumberAlignment
    {
        /// <summary>
        /// Align Left.
        /// </summary>
        Left = 0,
        /// <summary>
        /// Align Center.
        /// </summary>
        Center = 1,
        /// <summary>
        /// Align Right.
        /// </summary>
        Right = 2
    }

    /// <summary>
    /// Defines the LevelSuffixCharacter enumeration.
    /// </summary>
    public enum LevelSuffixCharacter
    {
        /// <summary>
        /// Tab Between Numbering and Text.
        /// </summary>
        Tab = 0,
        /// <summary>
        /// Space Between Numbering and Text.
        /// </summary>
        Space = 1,
        /// <summary>
        /// Nothing Between Numbering and Text.
        /// </summary>
        Nothing = 2
    }

    /// <summary>
    /// Defines the table cell vertical alignment enumerations.
    /// </summary>
    public enum TableCellVerticalAlignment
    {
        /// <summary>
        /// top.
        /// </summary>
        Top = 0,
        /// <summary>
        /// center.
        /// </summary>
        Center = 1,
        /// <summary>
        /// bottom.
        /// </summary>
        Bottom = 2
    }

    /// <summary>
    /// Defines the page borders position enumeration.
    /// </summary>
    public enum PageBordersPosition
    {
        /// <summary>
        /// Page Border Is Positioned Relative to Page Edges. 
        /// </summary>
        Page = 0,
        /// <summary>
        /// Page Border Is Positioned Relative to Text Extents. 
        /// </summary>
        Text = 1
    }

    /// <summary>
    /// Defines the AutoFitMethod enumeration.
    /// </summary>
    public enum AutoFitMethod
    {
        /// <summary>
        /// Resize the table columns to be the same with as contents.
        /// </summary>
        AutoFitContents = 0,
        /// <summary>
        /// Resize the table columns to stretch across the page.
        /// </summary>
        AutoFitWindow = 1,
        /// <summary>
        /// Resize the table columns to fixed width.
        /// </summary>
        FixedColumnWidth = 2
    }

    /// <summary>
    /// Defines the CellWidthType enumeration.
    /// </summary>
    public enum CellWidthType
    {
        /// <summary>
        /// Automatically Determined Width. 
        /// </summary>
        Auto = 0,
        /// <summary>
        /// Width in Percent.
        /// </summary>
        Percent = 1,
        /// <summary>
        /// Width in Point. 
        /// </summary>
        Point = 2
    }

    /// <summary>
    /// Defines the TableRowAlignment enumeration.
    /// </summary>
    public enum TableRowAlignment
    {
        /// <summary>
        /// Left.
        /// </summary>
        Left = 0,
        /// <summary>
        /// Center.
        /// </summary>
        Center = 1,
        /// <summary>
        /// Right.
        /// </summary>
        Right = 2
    }
}
