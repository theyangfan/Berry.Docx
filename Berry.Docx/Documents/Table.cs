using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Collections;
using Berry.Docx.Field;
using Berry.Docx.Utils;

namespace Berry.Docx.Documents
{
    public class Table : DocumentElement
    {
        private Document _doc;
        private W.Table _table;

        public Table(Document doc, int rowCnt, int columnCnt)
            :this(doc, TableGenerator.GenerateTable(rowCnt, columnCnt))
        {
        }

        internal Table(Document doc, W.Table table):base(doc, table)
        {
            _doc = doc;
            _table = table;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Table;

        public override DocumentObjectCollection ChildObjects => Rows;

        public TableRowCollection Rows => new TableRowCollection(_table, TableRowsPrivate());
        private IEnumerable<TableRow> TableRowsPrivate()
        {
            foreach(W.TableRow row in _table.Elements<W.TableRow>())
            {
                yield return new TableRow(_doc, row);
            }
        }

        /// <summary>
        /// 从父类集合中移除当前表格
        /// </summary>
        internal void Remove()
        {
            if (_table != null) _table.Remove();
        }

    }
}
