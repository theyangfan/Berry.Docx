using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;
using Berry.Docx.Formatting;

namespace Berry.Docx.Collections
{
    public class ColumnWidthCollection : IEnumerable<ColumnWidth>
    {
        #region Private Members
        private readonly Document _doc;
        private W.TableGrid _tableGrid;
        #endregion

        #region Constructor
        internal ColumnWidthCollection(Document doc, Table table)
        {
            _doc = doc;
            _tableGrid = table.XElement.GetFirstChild<W.TableGrid>();
        }
        #endregion

        #region Public Properties
        public ColumnWidth this[int index] => Columns().ElementAt(index);
        #endregion


        #region Public Methods

        public IEnumerator<ColumnWidth> GetEnumerator()
        {
            return Columns().GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        #endregion

        #region Private Methods
        private IEnumerable<ColumnWidth> Columns()
        {
            if (_tableGrid == null) yield break;
            foreach (var column in _tableGrid.Elements<W.GridColumn>())
            {
                yield return new ColumnWidth(_doc, column);
            }
        }
        #endregion
    }
}
