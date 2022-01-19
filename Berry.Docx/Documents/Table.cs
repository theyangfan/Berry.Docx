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
            :this(doc, TableGenerator.Generate(rowCnt, columnCnt)) 
        {
        }

        internal Table(Document doc, W.Table table):base(doc, table)
        {
            _doc = doc;
            _table = table;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Table;

        public override DocumentObjectCollection ChildObjects
        {
            get
            {
                return new DocumentElementCollection(_table, ChildObjectsPrivate());
            }
        }

        public int RowCount => _table.Elements<W.TableRow>().Count();

        private IEnumerable<DocumentElement> ChildObjectsPrivate()
        {
            foreach (O.OpenXmlElement ele in _table.ChildElements)
            {
                if (ele.GetType() == typeof(W.Paragraph))
                    yield return new Paragraph(_doc, ele as W.Paragraph);
                else if (ele.GetType() == typeof(W.Run))
                    yield return new TextRange(_doc, ele as W.Run);
            }
        }

    }
}
