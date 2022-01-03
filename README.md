# Berry.Docx

Berry.Docx 是一款用于读写 Word 2007+ (.docx)文档的.NET 库。开发该库的主要目的是为了提供简便，友好的接口来处理底层的 [OpenXML](https://github.com/OfficeDev/Open-XML-SDK) API。



# 示例

Berry.Docx 支持读写 Word 文档，无需安装 MS Word 程序。

以下是一个设置段落格式的典型例子：

```c#
using Berry.Docx;
using Berry.Docx.Documents;

Document doc = new Document("test.docx");
Paragraph p = doc.Paragraphs[0];
//字体
p.CharacterFormat.FontCN = "微软雅黑";
// 字号
p.CharacterFormat.FontSize = 16;
// 加粗
p.CharacterFormat.Bold = true;
// 对齐方式
p.Format.Justification = JustificationType.Left;
// 大纲级别
p.Format.OutlineLevel = OutlineLevelType.Level1;
// 左侧缩进
p.Format.LeftCharsIndent = 2;
// 段前行距
p.Format.BeforeLinesSpacing = 1;
//行距
p.Format.LineSpacing = 12;
p.Format.LineSpacingRule = LineSpacingRule.Multiple;

doc.Save();
doc.Close();
```



# 依赖

本项目基于.NET Framework 4.7.2 开发，依赖的 [OpenXML SDK](https://github.com/OfficeDev/Open-XML-SDK) 版本为 v2.15.0 。



# 更新日志

##### v1.0.0（2022-01-03）

- 支持创建、打开 .docx 文档；
- 支持读取、设置段落的常规格式(字体、对齐方式、缩进、间距等)。