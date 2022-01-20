using Berry.Docx.Collections;
using System;
using System.Collections.Generic;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    public class TableRow : DocumentElement
    {
        private Document _doc;
        private W.TableRow _row;
        public TableRow(Document doc)
            :this(doc, new W.TableRow())
        {

        }

        internal TableRow(Document doc, W.TableRow row):base(doc, row)
        {
            _doc = doc;
            _row = row;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.TableRow;

        public override DocumentObjectCollection ChildObjects => Cells;

        public TableCellCollection Cells => new TableCellCollection(_row, GetTableCells());

        private IEnumerable<TableCell> GetTableCells()
        {
            foreach(W.TableCell cell in _row.Elements<W.TableCell>())
            {
                yield return new TableCell(_doc, cell);
            }
        }
    }
}
