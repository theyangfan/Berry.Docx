using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Represent the <see cref="int"/> value.
    /// </summary>
    internal class IntegerValue : SimpleValue<int>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="IntegerValue"/> class.
        /// </summary>
        public IntegerValue() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="IntegerValue"/> class using the supplied <see cref="int"/> value.
        /// </summary>
        /// <param name="value">The <see cref="int"/> value.</param>
        public IntegerValue(int value) : base(value) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="IntegerValue"/> class by deep copying
        /// the supplied <see cref="IntegerValue"/> class.
        /// </summary>
        /// <param name="source">The source <see cref="IntegerValue"/> class.</param>
        public IntegerValue(IntegerValue source) : base(source) { }

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
