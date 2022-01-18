using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Utils
{
    internal class TableGenerator
    {
        private static int TableWidth = 8296;
        public static Table Generate(int rowCnt, int columnCnt)
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
            TableStyle tableStyle1 = new TableStyle() { Val = "a3" };
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableLook tableLook1 = new TableLook() { FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableLook1);

            table.Append(tableProperties1);
        }

        private static void SetTableGrid(Table table, int columnCnt)
        {
            int averageWidth = TableWidth / columnCnt;

            TableGrid tableGrid1 = new TableGrid();
            for (int i = 1; i <= columnCnt; i++)
            {
                GridColumn gridColumn = new GridColumn();
                if (i < columnCnt)
                    gridColumn.Width = averageWidth.ToString();
                else
                    gridColumn.Width = (TableWidth - (averageWidth * (columnCnt - 1))).ToString();

                tableGrid1.Append(gridColumn);
            }
            table.Append(tableGrid1);
        }

        private static void SetTableData(Table table, int rowCnt, int columnCnt)
        {
            int averageWidth = TableWidth / columnCnt;
            for(int r = 1; r <= rowCnt; r++)
            {
                TableRow row = new TableRow();
                for(int c = 1; c <= columnCnt; c++)
                {
                    TableCell cell = new TableCell();
                    TableCellProperties tableCellProperties = new TableCellProperties();
                    TableCellWidth tableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Dxa };
                    if (c < columnCnt)
                        tableCellWidth.Width = averageWidth.ToString();
                    else
                        tableCellWidth.Width = (TableWidth - (averageWidth * (columnCnt - 1))).ToString();

                    tableCellProperties.Append(tableCellWidth);
                    Paragraph paragraph = new Paragraph();

                    cell.Append(tableCellProperties);
                    cell.Append(paragraph);

                    row.Append(cell);
                }
                table.Append(row);
            }
        }


    }
}
