using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Represent the <see cref="Enum"/> value.
    /// </summary>
    internal class EnumValue<T> : SimpleValue<T> where T : struct
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="EnumValue{T}"/> class.
        /// </summary>
        public EnumValue() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="EnumValue{T}"/> class using the supplied <see cref="Enum"/> value.
        /// </summary>
        /// <param name="value">The <see cref="Enum"/> value.</param>
        public EnumValue(T value) : base(value) { }


        /// <summary>
        /// Initializes a new instance of the <see cref="EnumValue{T}"/> class by deep copying
        /// the supplied <see cref="EnumValue{T}"/> class.
        /// </summary>
        /// <param name="source">The source <see cref="EnumValue{T}"/> class.</param>
        public EnumValue(EnumValue<T> source) : base(source) { }

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
