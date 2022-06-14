using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Represent the <see cref="Enum"/> value.
    /// </summary>
    public class EnumValue<T> where T : struct
    {
        private T _value = default(T);
        /// <summary>
        /// Initializes a new instance of the <see cref="EnumValue{T}"/> class.
        /// </summary>
        public EnumValue() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="EnumValue{T}"/> class using the supplied <see cref="Enum"/> value.
        /// </summary>
        /// <param name="value">The <see cref="Enum"/> value.</param>
        public EnumValue(T value)
        {
            _value = value;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EnumValue{T}"/> class by deep copying
        /// the supplied <see cref="EnumValue{T}"/> class.
        /// </summary>
        /// <param name="source">The source <see cref="EnumValue{T}"/> class.</param>
        public EnumValue(EnumValue<T> source)
        {
            _value = source.Val;
        }

        /// <summary>
        /// Gets or sets the inner <see cref="Enum"/> value.
        /// </summary>
        public T Val
        {
            get => _value;
            set => _value = value;
        }

        /// <summary>
        /// Implicitly converting the <see cref="EnumValue{T}"/> value to <see cref="Enum"/> value.
        /// </summary>
        /// <param name="value">The <see cref="EnumValue{T}"/> value.</param>
        public static implicit operator T(EnumValue<T> value)
        {
            return value.Val;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EnumValue{T}"/> class by implicitly
        /// converting the supplied <see cref="Enum"/> value.
        /// </summary>
        /// <param name="value">The <see cref="Enum"/> value.</param>
        public static implicit operator EnumValue<T>(T value)
        {
            return new EnumValue<T>(value);
        }
    }
}
