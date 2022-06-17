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
        /// Undefined rule.
        /// </summary>
        None = -1,
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
        Normal = 0,
        Heading1 = 1,
        Heading2 = 2,
        Heading3 = 3,
        Heading4 = 4,
        Heading5 = 5,
        Heading6 = 6,
        Heading7 = 7,
        Heading8 = 8,
        Heading9 = 9
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
}
