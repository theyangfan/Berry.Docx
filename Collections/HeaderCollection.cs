using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Berry.Docx.Documents;

namespace Berry.Docx.Collections
{
    

    public class HeaderCollection
    {
        private List<Header> _headers;

        public HeaderCollection(List<Header> headers)
        {
            _headers = headers;
        }

        public Header this[int index]
        {
            get 
            {
                return _headers[index];
            }
        }


        public IEnumerator GetEnumerator()
        {
            return new HeaderEnumerator(_headers);
        }
    }


    public class HeaderEnumerator : IEnumerator
    {
        private List<Header> _headers;
        int _position = -1;
        public HeaderEnumerator(List<Header> headers)
        {
            _headers = headers;
        }

        public object Current
        {
            get
            {
                if (_position == -1)
                    throw new InvalidOperationException();
                if (_position >= _headers.Count)
                    throw new InvalidOperationException();
                return _headers[_position];
            }
        }

        public bool MoveNext()
        {
            if (_position < _headers.Count - 1)
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
