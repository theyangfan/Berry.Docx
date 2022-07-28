using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Berry.Docx.Field;

namespace Berry.Docx.Collections
{
    internal class FieldCodeCollection : IEnumerable
    {
        private List<FieldCode> _fieldcodes;
        public FieldCodeCollection(List<FieldCode> fieldcodes)
        {
            _fieldcodes = fieldcodes;
        }

        public FieldCode this[int index] { get => _fieldcodes[index]; }

        public int Count { get => _fieldcodes.Count; }

        public IEnumerator GetEnumerator()
        {
            return new FieldCodeEnumerator(_fieldcodes);
        }

    }

    internal class FieldCodeEnumerator : IEnumerator
    {
        private List<FieldCode> _fieldcodes;
        int _position = -1;
        public FieldCodeEnumerator(List<FieldCode> fieldcodes)
        {
            _fieldcodes = fieldcodes;
        }
        public object Current
        {
            get
            {
                if (_position == -1)
                    throw new InvalidOperationException();
                if (_position >= _fieldcodes.Count)
                    throw new InvalidOperationException();
                return _fieldcodes[_position];
            }
        }

        public bool MoveNext()
        {
            if (_position < _fieldcodes.Count - 1)
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
