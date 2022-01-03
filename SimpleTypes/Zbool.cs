using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx
{
    public class Zbool
    {
        private bool _value = false;

        public Zbool() { }

        public Zbool(bool value)
        {
            _value = value;
        }

        public Zbool(Zbool val)
        {
            _value = val.Val;
        }

        public bool Val
        {
            get => _value;
            set => _value = value;
        }

        public static implicit operator bool(Zbool value)
        {
            return value.Val;
        }

        public static implicit operator Zbool(bool value)
        {
            return new Zbool(value);
        }
    }
}
