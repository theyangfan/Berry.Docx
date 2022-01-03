using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Berry.Docx.Documents;
using System.Collections;

namespace Berry.Docx.Collections
{
    public class FooterCollection
    {
        private List<Footer> _footers;

        public FooterCollection(List<Footer> footers)
        {
            _footers = footers;
        }

        public Footer this[int index]
        {
            get
            {
                return _footers[index];
            }
        }

        public IEnumerator GetEnumerator()
        {
            return new FooterEnumerator(_footers);
        }
    }

    public class FooterEnumerator : IEnumerator
    {
        private List<Footer> _footer;
        int _position = -1;
        public FooterEnumerator(List<Footer> footer)
        {
            _footer = footer;
        }

        public object Current
        {
            get
            {
                if (_position == -1)
                    throw new InvalidOperationException();
                if (_position >= _footer.Count)
                    throw new InvalidOperationException();
                return _footer[_position];
            }
        }

        public bool MoveNext()
        {
            if (_position < _footer.Count - 1)
            {
                _position++;
                return true;
            }
            else
            {
                return false;
            }
        }

        public void Reset()
        {
            _position = -1;
        }
    }




}
