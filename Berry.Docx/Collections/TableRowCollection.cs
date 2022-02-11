using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using O = DocumentFormat.OpenXml;

using Berry.Docx.Documents;

namespace Berry.Docx.Collections
{
    public class TableRowCollection : DocumentElementCollection
    {
        private IEnumerable<TableRow> _rows;
        internal TableRowCollection(O.OpenXmlElement owner, IEnumerable<TableRow> rows)
            : base(owner, rows)
        {
            _rows = rows;
        }

        public new TableRow this[int index] => _rows.ElementAt(index);

        public TableRow Last()
        {
            return _rows.Last();
        }
    }
}
