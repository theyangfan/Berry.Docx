using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using Berry.Docx.Formatting;

namespace Berry.Docx.Collections
{
    public class ListLevelCollection : IEnumerable<ListLevel>
    {
        private readonly IEnumerable<ListLevel> _levels;
        internal ListLevelCollection(IEnumerable<ListLevel> levels)
        {
            _levels = levels;
        }

        public ListLevel this[int index] => _levels.ElementAt(index);

        public int Count => _levels.Count();

        public IEnumerator<ListLevel> GetEnumerator()
        {
            return _levels.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
