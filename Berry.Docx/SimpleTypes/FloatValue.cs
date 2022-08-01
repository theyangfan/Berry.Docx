using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Represent the <see cref="float"/> value.
    /// </summary>
    internal class FloatValue : SimpleValue<float>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FloatValue"/> class.
        /// </summary>
        public FloatValue() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="FloatValue"/> class using the supplied <see cref="float"/> value.
        /// </summary>
        /// <param name="value">The <see cref="float"/> value.</param>
        public FloatValue(float value) : base(value) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="FloatValue"/> class by deep copying
        /// the supplied <see cref="FloatValue"/> class.
        /// </summary>
        /// <param name="source">The source <see cref="FloatValue"/> class.</param>
        public FloatValue(FloatValue source) : base(source) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="FloatValue"/> class by implicitly
        /// converting the supplied <see cref="float"/> value.
        /// </summary>
        /// <param name="value">The <see cref="float"/> value.</param>
        public static implicit operator FloatValue(float value)
        {
            return new FloatValue(value);
        }
    }
}
