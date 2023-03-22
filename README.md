# Berry.Docx

[![Downloads](https://img.shields.io/nuget/dt/Berry.Docx.svg)](https://www.nuget.org/packages/Berry.Docx)

根据 Office Open XML 格式规范(ECMA 375 第二版) 开发的一款用于读写 Word 2007+ (.docx) 文档的 .NET 库，基于 OpenXML SDK。作为一个 .NET 新手，建立此项目的主要目的是锻炼自己的 API 设计能力，目前项目仍在建设中...

A .NET library for reading and writing Word 2007+ (.docx) files developed according to the Office OpenXML Format specification (ECMA-375, second edition), based on the OpenXML SDK. As a .NET novice, the main purpose of this project is to exercise my API design ability, which is still under construction...

<br/>

# 程序包（Packages）

Berry.Docx 的 NuGet 软件包发布在NuGet.org上:

*The release NuGet packages for Berry.Docx are on NuGet.org:*

| Package    | Download                                                                                             |
| ---------- | ---------------------------------------------------------------------------------------------------- |
| Berry.Docx | [![NuGet](https://img.shields.io/nuget/v/Berry.Docx.svg)](https://www.nuget.org/packages/Berry.Docx) |

## 通过 NuGet 安装（Install via NuGet）

如果你想在项目中使用 Berry.Docx，最简单的方法就是通过 NuGet 包管理器安装。

用 Visual Studio 打开自己的项目，在项目上右键选择【管理 NuGet 程序包】选项，在浏览输入框中输入“Berry.Docx”，如下所示：

*If you want to include Berry.Docx in your project, you can install it directly from NuGet.*

*Open your project in Visual Studio, right-click the solution and select  **Manager NuGet Packages** , then enter "Berry.Docx" in the Browse input box, as follows:*

![image](https://github.com/theyangfan/Berry.Docx/blob/main/nuget_package_manager.png)

选择安装即可。

或者在 Visual Studio 程序包管理器控制台运行以下命令来安装：

*Select and install.*

*Or you could run the following command in the Package Manager Console to install it.*

```sh
PM> Install-Package Berry.Docx
```

<br/>

# 示例（Examples）

下面的示例演示如何新增一个文档并添加一个格式为“微软雅黑、小四、居中”的段落，以及一个3行3列的表格。

*The following example shows how to create a new document file, and add a new paragraph with "Times New Roman font, 14 point, Center justification" format,and a 3x3 size table.*

```c#
using Berry.Docx;
using Berry.Docx.Documents;

namespace Example
{
    class Example
    {
        static void Main() 
        {
            // 新建一个名为“example.docx”的文档 (Create a new word document called "example.docx")
            using (Document doc = new Document("example.docx"))
            {
                // 新建一个段落 (Create a new paragraph)
                Paragraph p1 = doc.CreateParagraph();
                p1.Text = "这是一个段落。This is a paragraph.";
                foreach(TextRange tr in p1.ChildItems.OfType<TextRange>())
                {
                    tr.CharacterFormat.FontNameAscii = "Times New Roman";
                    tr.CharacterFormat.FontNameEastAsia = "微软雅黑";
                    tr.CharacterFormat.FontSize = 14;
                }
                p1.Format.Justification = JustificationType.Center;
                // 新建一个表格 (Create a new table)
                Table tbl1 = doc.CreateTable(3, 3);
                tbl1.Rows[0].Cells[1].Paragraphs[0].Text = "1st Column";
                tbl1.Rows[0].Cells[2].Paragraphs[0].Text = "2nd Column";
                tbl1.Rows[1].Cells[0].Paragraphs[0].Text = "1st Row";
                tbl1.Rows[2].Cells[0].Paragraphs[0].Text = "2nd Row";
                // 添加到文档中 (Add to the document)
                doc.Sections[0].ChildObjects.Add(p1);
                doc.Sections[0].ChildObjects.Add(tbl1);
                // 保存 (Save)
                doc.Save();
            }
        } 
    }
}
```

<br/>

# 主要功能（Main Features）

| Features                                                                                                                  |
| ------------------------------------------------------------------------------------------------------------------------- |
| 操作段落和字符\|*Manipulates paragraphs and characters*                                                                          |
| 操作表格及其行和单元格\|*Manipulates table and it's rows and cells*                                                                  |
| 读写字符格式(中文字体，西文字体，字号，加粗，斜体等)\|*Read-write character format (FontNameEastAsia, FontNameAscii, FontSize, Bold, Italic etc.)* |
| 读写段落格式(对齐方式, 大纲级别, 缩进, 间距等.)\|*Read-write paragraph format (Justification, OutlineLevel, Indentation, Spacing etc.)*      |
| 读写段落、表格、字符样式\|*Read-write paragraph, table, character style*                                                              |
| 插入分节符，分页符和手动换行符\|*Inserts section break，page break and line break*                                                        |
| 添加批注\|*Appends comments*                                                                                                  |
| 操作页眉页脚\|*Manipulates header and footers*                                                                                  |
| 读写页面设置\|*Read-write page setup*                                                                                           |
| 读写列表样式\|*Read-write list style*                                                                                           |
| 查找文本\|*Find text*                                                                                                         |
| 读写表格格式 \| *Read-write table formats*                                                                                      |
| 向段落中添加图片、分隔符和制表位 \| *Append picture , break and tab to paragraph*                                                         |

<br/>

# 文档（Documentation）

- [API Documentation](https://theyangfan.github.io/Berry.Docx)

<br/>

# 更新日志（Release History）

### v1.3.3 (2023-03-22)

- 支持向段落中添加图片、分隔符和制表位 (*Support append picture , break and tab to paragraph*)；

- 支持更多的格式 (*Support more formats*)；

- 修复了在解析无效超链接时抛出异常的问题 (*Fixed an exception being thrown when parsing malformed hyperlinks*)。

### v1.3.2 (2022-08-26)

- 支持读写表格格式 (*Support read-write table formats*)。

### v1.3.1 (2022-08-02)

- 修复 `TextMatch.GetAsOneRange()` 方法抛出异常的问题 (Fixed the bug that  `TextMatch.GetAsOneRange()` throw an exception)。

### v1.3.0 (2022-08-01)

- 支持更多字符和段落格式 (*Support more character and paragraph format*)；
- 支持读写页面设置 (*Support read-write page setup*)；
- 支持读写段落、字符和表格样式(*Support read-write paragraph, table and character style*)；
- 支持读写列表样式 (*Support read-write list style*)；
- 支持查找文本 (*Support find text*)。

### v1.2.0 (2022-03-22)

- 支持操作页眉和页脚 （*Supports manipulating headers and footers*)。

### v1.1.0 (2022-03-06)

- 支持更多字符和段落高级格式 (*Supports more character and paragraph advanced formats*)；
- 支持插入分节符 (*Supports insert section break*)；
- 支持插入批注 (*Supports append paragraph comments*)。

### v1.0.1 (2022-02-11)

- 支持新增/删除段落 (*Supports add/remove paragraphs*)；

- 支持新增/删除表格，以及在表格中插入行，列，单元格和段落 (*Supports add/remove tables, and add/insert rows, columns, cells, paragraphs to the table*)。

### v1.0.0（2022-01-03）

- 支持创建、打开 DOCX 文档 (*Supports create & open DOCX files*)；
- 支持读取、设置段落的常规格式 (字体、对齐方式、缩进、间距) (*Supports read and change the normal format (Font、Justification、Indentation、Spacing) of paragraph*)。

<br/>

# 下版本计划（Next Version Plan）

- 支持读写域代码 (*Support read-write field codes*)
- 支持更多图形属性 (*Support more shape properties*)
