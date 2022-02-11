**English** | [简体中文](./README_CN.md)

<br/>

# Berry.Docx

Berry.Docx is a .NET library for reading, manipulating and writing Word 2007+ (.docx) files without the Word application. It aims to provide an intuitive, full and user-friendly interface to dealing with the underlying [OpenXML](https://github.com/OfficeDev/Open-XML-SDK) API.

<br/>

# Install via NuGet

If you want to include Berry.Docx in your project, you can install it directly from NuGet.

Run the following command in the Package Manager Console to install it.

```sh
PM> Install-Package Berry.Docx
```

<br/>

# Examples

The following example shows how to create a new document file, and add a new paragraph with "Times New Roman font, 14 point, Center justification" format,and a 3x3 size table.

```c#
using Berry.Docx;
using Berry.Docx.Documents;

namespace Example
{
    class Example
    {
        static void Main() 
        {
			// Create a new word document called “example.docx”
            Document doc = new Document("example.docx");
			// Create a new paragraph
            Paragraph p1 = doc.CreateParagraph();
            p1.Text = "This is a paragraph.";
            p1.CharacterFormat.FontEN = "Times New Roman";
            p1.CharacterFormat.FontSize = 14;
            p1.Format.Justification = JustificationType.Center;
			// Create a new table
            Table tbl1 = doc.CreateTable(3, 3);
            tbl1.Rows[0].Cells[1].Paragraphs[0].Text = "1st Column";
            tbl1.Rows[0].Cells[2].Paragraphs[0].Text = "2nd Column";
            tbl1.Rows[1].Cells[0].Paragraphs[0].Text = "1st Row";
            tbl1.Rows[2].Cells[0].Paragraphs[0].Text = "2nd Row";
			// Add to the document
            doc.Sections[0].Range.ChildObjects.Add(p1);
            doc.Sections[0].Range.ChildObjects.Add(tbl1);
			// Save and close
            doc.Save();
            doc.Close();
        } 
    }
}
```

<br/>

# Main Features

| Features                                                     |
| ------------------------------------------------------------ |
| Open existing DOCX files Or Create new DOCX files            |
| Get sections                                                 |
| Get paragraphs of section Or Append/Insert new paragraphs    |
| Get characters of paragraph Or Append/Insert new characters  |
| Get/Set character format(FontCN, FontEN, FontSize, Bold, Italic) |
| Get/Set paragraph format(Justification, OutlineLevel, Indentation, Spacing etc.) |
| Get paragraph style                                          |
| Get tables of section Or Append/Insert new tables            |
| Get table rows and cells                                     |
| Insert Rows/Columns around table cells                       |
| Get/Set table cell paragraphs                                |

<br/>

# Documentation

- [Examples](https://theyangfan.github.io/Berry.Docx/en-US/examples/paragraph/index.html)
- [APIs References](https://theyangfan.github.io/Berry.Docx/en-US/api/index.html)

<br/>

# Release History

### v1.0.1 (2022-02-11)

#### Added

- Supports add/remove paragraphs;

- Supports add/remove tables, and add/insert rows, cloumns, cells, paragraphs to the table.

### v1.0.0（2022-01-03）

#### Added

- Supports create & open DOCX files；
- Supports read and change the normal format(Font、Justification、Indentation、Spacing) of paragraph.

