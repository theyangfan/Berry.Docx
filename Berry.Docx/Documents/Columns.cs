using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// Represents the page setup columns.
    /// </summary>
    public class Columns : IEnumerable
    {
        #region Private Members
        private readonly Document _doc;
        private readonly Section _sect;
        private readonly W.Columns _columns;
        #endregion

        #region Constructors
        internal Columns(Document doc, Section section, W.Columns columns)
        {
            _doc = doc;
            _sect = section;
            _columns = columns;
        }
        #endregion

        #region Public Properties

        
        public Column this[int index] => GetColumns().ElementAt(index);

        public bool EqualColumnWidth
        {
            get
            {
                return _columns.EqualWidth ?? true;
            }
            set
            {
                _columns.EqualWidth = value;
            }
        }

        public int EqualWidthColumnsCount
        {
            get
            {
                return _columns.ColumnCount ?? 1;
            }
            set
            {
                _columns.ColumnCount = (short)value;
            }
        }

        public float EqualWidthColumnsSpacing
        {
            get
            {
                if (_columns.Space == null) return 0;
                return (_columns.Space.ToString().ToInt() / 20.0F).Round(2);
            }
            set
            {
                _columns.Space = (value * 20).Round(0).ToString();
            }
        }

        public bool LineBetweenColumns
        {
            get
            {
                return _columns.Separator ?? false;
            }
            set
            {
                _columns.Separator = value;
            }
        }
        #endregion

        #region Public Methods

        
        public int Count()
        {
            return GetColumns().Count();
        }
        public void Add(Column column)
        {
            _columns.Append(column.XElement);
        }

        public void Clear()
        {
            _columns.ColumnCount = null;
            _columns.RemoveAllChildren();
        }

        public IEnumerator GetEnumerator()
        {
            return GetColumns().GetEnumerator();
        }
        #endregion

        #region Private Methods
        private IEnumerable<Column> GetColumns()
        {
            if (EqualColumnWidth)
            {
                for (int i = 1; i <= EqualWidthColumnsCount; i++)
                    yield return new Column(_doc) { Spacing = EqualWidthColumnsSpacing };
            }
            else
            {
                foreach (W.Column column in _columns.Elements<W.Column>())
                    yield return new Column(_doc, column);
            }
        }
        #endregion
    }
}
