using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx
{
    internal class TableGenerator
    {
        public static Table GenerateTable(int rowCnt, int columnCnt)
        {
            Table table1 = new Table();

            SetTableProperties(table1);

            SetTableGrid(table1, columnCnt);

            SetTableData(table1, rowCnt, columnCnt);

            return table1;
        }

        private static void SetTableProperties(Table table)
        {
            TableProperties tableProperties1 = new TableProperties();

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);

            TableLook tableLook1 = new TableLook() { FirstRow = false, LastRow = false, FirstColumn = false, LastColumn = false, NoHorizontalBand = true, NoVerticalBand = true };

            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLook1);

            table.Append(tableProperties1);
        }

        private static void SetTableGrid(Table table, int columnCnt)
        {
            TableGrid tableGrid1 = new TableGrid();
            table.Append(tableGrid1);
        }

        private static void SetTableData(Table table, int rowCnt, int columnCnt)
        {
            for(int r = 1; r <= rowCnt; r++)
            {
                TableRow row = new TableRow();

                TableRowProperties tableRowProperties1 = new TableRowProperties();
                row.Append(tableRowProperties1);

                for (int c = 1; c <= columnCnt; c++)
                {
                    TableCell tableCell1 = new TableCell();

                    TableCellProperties tableCellProperties1 = new TableCellProperties();

                    TableCellBorders tableCellBorders1 = new TableCellBorders();
                    TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                    LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                    BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                    RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                    InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                    InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                    tableCellBorders1.Append(topBorder1);
                    tableCellBorders1.Append(leftBorder1);
                    tableCellBorders1.Append(bottomBorder1);
                    tableCellBorders1.Append(rightBorder1);
                    tableCellBorders1.Append(insideHorizontalBorder1);
                    tableCellBorders1.Append(insideVerticalBorder1);
                    Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                    tableCellProperties1.Append(tableCellBorders1);
                    tableCellProperties1.Append(shading1);

                    Paragraph paragraph1 = new Paragraph();
                    ParagraphProperties paragraphProperties1 = new ParagraphProperties();

                    paragraph1.Append(paragraphProperties1);

                    tableCell1.Append(tableCellProperties1);
                    tableCell1.Append(paragraph1);

                    row.Append(tableCell1);
                }
                table.Append(row);

            }
        }


    }
}
