# Berry.Docx.Visual

[![Downloads](https://img.shields.io/nuget/dt/Berry.Docx.Visual.svg)](https://www.nuget.org/packages/Berry.Docx.Visual)

一款用于构建 DOCX 文档在显示时的组成结构的 .NET 库。通过此项目，我们可以访问文档中的页面、表格、段落以及其中的字符和图片等，以及这些内容的尺寸和边距等格式。

A.NET library for building the structure of a DOCX document when it is displayed. Through this project, we can access the pages, tables, paragraphs, characters and images in the document, as well as the format of the dimensions and margins of these contents.



本项目的计算结果是参照文档在 Microsoft Office Word 2019 的显示效果而生成的。由于无从得知 Word 的渲染方案，所以不保证结果与其一致。

The calculation result of this project is generated by referring to the document display effect in Microsoft Office Word 2019. Since Word's rendering scheme is not known, there is no guarantee that the results will be consistent. 



如果你想快速了解显示效果，可以查看 [Berry.DocxViewer](https://github.com/theyangfan/Berry.DocxViewer) 项目，这是一款 WPF 控件库，支持浏览 DOCX 文档。

If you want to get a quick look at the display effect, please check [Berry. DocxViewer](https://github.com/theyangfan/Berry.DocxViewer) project, that's a WPF control library, supports browsing DOCX documents.

## 程序包（Packages）

Berry.Docx.Visual 的 NuGet 软件包发布在NuGet.org上:

*The release NuGet packages for Berry.Docx.Visual are on NuGet.org:*

| Package           | Download                                                                                                           |
| ----------------- | ------------------------------------------------------------------------------------------------------------------ |
| Berry.Docx.Visual | [![NuGet](https://img.shields.io/nuget/v/Berry.Docx.Visual.svg)](https://www.nuget.org/packages/Berry.Docx.Visual) |



## 示例（Examples）

下面的示例演示如何读取文档中的内容：

The following example shows how to read the document contents:

```c#
using Berry.Docx.Visual;
using Berry.Docx.Visual.Documents;
using Berry.Docx.Visual.Field;

using (Berry.Docx.Document doc = new Berry.Docx.Document("example.docx"))
{
	Document visualDoc = new Document(doc);
	// get first page
	var page1 = visualDoc.Pages[0];
	Console.WriteLine(page1.Width);
	Console.WriteLine(page1.Height);
	Console.WriteLine(page1.Padding);
	// get first item
	var item1 = page1.ChildItems[0];
	// paragraph
	if(item1 is Paragraph)
	{
		var paragraph1 = (Paragraph)item1;
		var pItem1 = paragraph1.Lines[0].ChildItems[0];
		if(pItem1 is Character)
		{
			var character1 = (Character)pItem1;
			Console.WriteLine(character1.Val);
		}
		else if(pItem1 is Picture)
		{
			var pic1 = (Picture)pItem1;
			Console.WriteLine(pic1.Width);
		}
	}
	// table
	else if(item1 is Table)
	{
		var table1 = (Table)item1;
		Console.WriteLine(table1.Width);
		Console.WriteLine(table1.Height);
		Console.WriteLine(table1.Cells.Count);
	}
}
```



## 更新日志（Release History）

### v1.0.2 (2023-06-16)

- 支持表格 (Supports tables)。

### v1.0.1 (2023-06-02)

- 支持嵌入式图片 (Supports inline pictures)。

### v1.0.0 (2023-05-29)

- 支持页面、段落和文本字符 (Supports pages、paragraphs and characters)。