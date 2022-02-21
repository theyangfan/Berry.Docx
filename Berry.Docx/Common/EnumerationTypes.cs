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
        BodyRange
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
        /// Character Grid Only.
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
        /// Undefined.
        /// </summary>
        None = -1,
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
        Right = 2,
        /// <summary>
        /// Justified.
        /// </summary>
        Both = 3,
        /// <summary>
        /// Distribute All Characters Equally.
        /// </summary>
        Distribute = 4
    }

    /// <summary>
    /// Defines the OutlineLevelType enumeration.
    /// </summary>
    public enum OutlineLevelType
    {
        /// <summary>
        /// Undefined level.
        /// </summary>
        None = -1,
        /// <summary>
        /// Level 1.
        /// </summary>
        Level1 = 0,
        /// <summary>
        /// Level 2.
        /// </summary>
        Level2 = 1,
        /// <summary>
        /// Level 3.
        /// </summary>
        Level3 = 2,
        /// <summary>
        /// Level 4.
        /// </summary>
        Level4 = 3,
        /// <summary>
        /// Level 5.
        /// </summary>
        Level5 = 4,
        /// <summary>
        /// Level 6.
        /// </summary>
        Level6 = 5,
        /// <summary>
        /// Level 7.
        /// </summary>
        Level7 = 6,
        /// <summary>
        /// Level 8.
        /// </summary>
        Level8 = 7,
        /// <summary>
        /// Level 9.
        /// </summary>
        Level9 = 8,
        /// <summary>
        /// Body Text.
        /// </summary>
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
}
