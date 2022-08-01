using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    internal abstract class SimpleValue<T> : IEquatable<SimpleValue<T>> where T : struct
    {
        private protected T _value = default(T);

        public SimpleValue(){}

        public SimpleValue(T value)
        {
            _value = value;
        }

        public SimpleValue(SimpleValue<T> source)
        {
            _value = source.Val;
        }

        /// <summary>
        /// Gets or sets the inner value.
        /// </summary>
        public T Val
        {
            get => _value;
            set => _value = value;
        }

        public static implicit operator T(SimpleValue<T> value)
        {
            return value.Val;
        }

        public static bool operator ==(SimpleValue<T> lhs, SimpleValue<T> rhs)
        {
            if (ReferenceEquals(lhs, rhs)) return true;
            if ((object)lhs == null || (object)rhs == null) return false;
            return lhs.Val.Equals(rhs.Val);
        }

        public static bool operator !=(SimpleValue<T> lhs, SimpleValue<T> rhs)
        {
            return !(lhs == rhs);
        }

        public bool Equals(SimpleValue<T> other)
        {
            return this == other;
        }

        public override string ToString()
        {
            return _value.ToString();
        }
    }
}
