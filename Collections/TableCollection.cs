using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Berry.Docx.Documents;

namespace Berry.Docx.Collections
{
    public class TableCollection : IEnumerable
    {
        private IEnumerable<Table> _tables;
        public TableCollection(IEnumerable<Table> tables)
        {
            _tables = tables;
        }

        public Table this[int index]
        {
            get
            {
                return _tables.ElementAt(index);
            }
        }

        /// <summary>
        /// 返回集合数量
        /// </summary>
        public int Count { get => _tables.Count(); }

        public IEnumerator GetEnumerator()
        {
            return _tables.GetEnumerator();
        }

    }
}
