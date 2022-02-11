using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using O = DocumentFormat.OpenXml;

using Berry.Docx.Documents;
namespace Berry.Docx.Collections
{
    public class TableCellCollection : DocumentElementCollection
    {
        private IEnumerable<TableCell> _cells;
        internal TableCellCollection(O.OpenXmlElement owner, IEnumerable<TableCell> cells)
            : base(owner, cells)
        {
            _cells = cells;
        }

        public new TableCell this[int index] => _cells.ElementAt(index);

        public TableCell Last()
        {
            return _cells.Last();
        }
    }
}
