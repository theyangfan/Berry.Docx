using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;

namespace Berry.Docx.Collections
{
    public class Columns : IEnumerable
    {
        private readonly W.Columns _columns;
        private readonly IEnumerable<Column> _column_list;
        
        internal Columns(Document doc, Section section, W.Columns columns, IEnumerable<Column> column_list)
        {
            _columns = columns;
            _column_list = column_list;
        }

        public void Add(Column column)
        {
            _columns.Append(column.XElement);
        }

        public bool LineBetweenColumns
        {
            get
            {
                if (_columns?.Separator == null) return false;
                return _columns.Separator;
            }
            set
            {
                if (_columns == null)
                {
                    _columns = new W.Columns();
                    _sect.XElement.AddChild(_columns);
                }
                _columns.Separator = value;
            }
        }

    }
}
