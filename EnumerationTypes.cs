using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// 文档对象类型
    /// </summary>
    public enum DocumentObjectType
    {
        /// <summary>
        /// 无效
        /// </summary>
        Invalid = -1,
        /// <summary>
        /// 段落
        /// </summary>
        Paragraph = 0,
        /// <summary>
        /// 表格
        /// </summary>
        Table = 1,
        /// <summary>
        /// 节
        /// </summary>
        Section = 2,
        /// <summary>
        /// 文本范围
        /// </summary>
        TextRange = 3,
        BodyRange
    }

    public enum LineSpacingRule 
    {
        None = -1,
        /// <summary>
        /// 行距大于等于 LineSpacing 属性的值
        /// </summary>
        AtLeast = 0,
        /// <summary>
        /// 行距固定，即使段落字体发生变化
        /// </summary>
        Exactly = 1,
        /// <summary>
        /// LineSpacing属性值为行的倍数，1行为12磅
        /// </summary>
        Multiple = 2
    }

    /// <summary>
    /// 多页
    /// </summary>
    public enum MultiPage
    {
        /// <summary>
        /// 普通
        /// </summary>
        Normal = 0,
        /// <summary>
        /// 对称页边距
        /// </summary>
        MirrorMargins = 1,
        /// <summary>
        /// 拼页
        /// </summary>
        PrintTwoOnOne = 2
    }

    public enum DocGridType
    {
        /// <summary>
        /// 无网格
        /// </summary>
        None = 0,
        /// <summary>
        /// 只指定行网格
        /// </summary>
        Lines = 1,
        /// <summary>
        /// 指定行和字符网格
        /// </summary>
        LinesAndChars = 2,
        /// <summary>
        /// 文字对齐字符网格
        /// </summary>
        SnapToChars = 3
    }
    /// <summary>
    /// 样式类型
    /// </summary>
    public enum StyleType
    {
        /// <summary>
        /// 段落样式
        /// </summary>
        Paragraph = 0,
        /// <summary>
        /// 字符样式
        /// </summary>
        Character = 1,
        /// <summary>
        /// 表格样式
        /// </summary>
        Table = 2,
        /// <summary>
        /// 编号样式
        /// </summary>
        Numbering = 3
    }

    public enum JustificationType
    {
        None = -1,
        Left = 0,
        Center = 1,
        Right = 2,
        Both = 3,
        Distribute = 4
    }

    public enum OutlineLevelType
    {
        None = -1,
        Level1 = 0,
        Level2 = 1,
        Level3 = 2,
        Level4 = 3,
        Level5 = 4,
        Level6 = 5,
        Level7 = 6,
        Level8 = 7,
        Level9 = 8,
        BodyText = 9
    }
}
