**简体中文** | [English](./README.md)

<br/>

# Berry.Docx

Berry.Docx 是一款用于读写 Word 2007+ (.docx) 文档的.NET 库，无需 Word 应用程序。 旨在提供简便，完整，友好的接口来封装底层的 [OpenXML](https://github.com/OfficeDev/Open-XML-SDK) API。

<br/>

# 通过 NuGet 安装

如果你想在项目中使用 Berry.Docx，最简单的方法就是通过 NuGet 包管理器安装。

请在 Visual Studio 程序包管理器控制台运行以下命令来安装：

```sh
PM> Install-Package Berry.Docx
```

<br/>

# 示例



下面的示例演示如何新增一个文档并添加一个格式为“微软雅黑、小四、居中”的段落，以及一个3行3列的表格。

```c#
using Berry.Docx;
using Berry.Docx.Documents;

namespace Example
{
    class Example
    {
        static void Main()
        {
			// 新建一个名为“示例.docx”的文档
            Document doc = new Document("示例.docx");
			// 新增一个段落
            Paragraph p1 = doc.CreateParagraph();
            p1.Text = "这是一个段落。";
            p1.CharacterFormat.FontCN = "微软雅黑";
            p1.CharacterFormat.FontSizeCN = "小四";
            p1.Format.Justification = JustificationType.Center;
			// 新增一个表格
            Table tbl1 = doc.CreateTable(3, 3);
            tbl1.Rows[0].Cells[1].Paragraphs[0].Text = "第1列";
            tbl1.Rows[0].Cells[2].Paragraphs[0].Text = "第2列";
            tbl1.Rows[1].Cells[0].Paragraphs[0].Text = "第1行";
            tbl1.Rows[2].Cells[0].Paragraphs[0].Text = "第2行";
			// 添加至文档中
            doc.Sections[0].Range.ChildObjects.Add(p1);
            doc.Sections[0].Range.ChildObjects.Add(tbl1);
			// 保存并关闭
            doc.Save();
            doc.Close();
        } 
    }
}
```

<br/>

# 主要功能

| 功能                                                    |
| ------------------------------------------------------- |
| 打开已有的 DOCX 文档或者创建新的 DOCX 文档              |
| 获取节                                                  |
| 获取节中的段落或者添加/插入新的段落                     |
| 获取段落中的字符或者添加/插入新的字符                   |
| 获取/设置字符格式(中文字体，西文字体，字号，加粗，斜体) |
| 获取/设置段落格式(对齐方式, 大纲级别, 缩进, 间距等.)    |
| 获取段落样式                                            |
| 获取节中的表格或者添加/插入新的表格                     |
| 在表格单元格四周插入新的行或列                          |
| 获取/设置表格单元格中的段落                             |

<br/>

# 文档

- [示例](https://theyangfan.github.io/Berry.Docx/zh-CN/examples/paragraph/index.html)
- [APIs 文档](https://theyangfan.github.io/Berry.Docx/zh-CN/api/index.html)

<br/>

# 更新日志

### v1.0.1 (2022-02-11)

#### 新增功能

- 支持新增/删除段落；
- 支持新增/删除表格，以及在表格中插入行，列，单元格和段落。

### v1.0.0（2022-01-03）

#### 新增功能

- 支持创建、打开 DOCX 文档；
- 支持读取、设置段落的常规格式(字体、对齐方式、缩进、间距)。

