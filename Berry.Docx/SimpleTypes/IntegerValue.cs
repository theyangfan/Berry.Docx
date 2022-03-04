using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Represent the <see cref="int"/> value.
    /// </summary>
    public class IntegerValue
    {
        private int _value = 0;
        /// <summary>
        /// Initializes a new instance of the <see cref="IntegerValue"/> class.
        /// </summary>
        public IntegerValue() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="IntegerValue"/> class using the supplied <see cref="int"/> value.
        /// </summary>
        /// <param name="value">The <see cref="int"/> value.</param>
        public IntegerValue(int value)
        {
            _value = value;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="IntegerValue"/> class by deep copying
        /// the supplied <see cref="IntegerValue"/> class.
        /// </summary>
        /// <param name="source">The source <see cref="IntegerValue"/> class.</param>
        public IntegerValue(IntegerValue source)
        {
            _value = source.Val;
        }

        /// <summary>
        /// Gets or sets the inner <see cref="int"/> value.
        /// </summary>
        public int Val
        {
            get => _value;
            set => _value = value;
        }

        /// <summary>
        /// Implicitly converting the <see cref="IntegerValue"/> value to <see cref="int"/> value.
        /// </summary>
        /// <param name="value">The <see cref="IntegerValue"/> value.</param>
        public static implicit operator int(IntegerValue value)
        {
            return value.Val;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="IntegerValue"/> class by implicitly
        /// converting the supplied <see cref="int"/> value.
        /// </summary>
        /// <param name="value">The <see cref="int"/> value.</param>
        public static implicit operator IntegerValue(int value)
        {
            return new IntegerValue(value);
        }
    }
}
