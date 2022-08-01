using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Represent the <see cref="bool"/> value.
    /// </summary>
    internal class BooleanValue : SimpleValue<bool>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="BooleanValue"/> class.
        /// </summary>
        public BooleanValue() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BooleanValue"/> class using the supplied <see cref="bool"/> value.
        /// </summary>
        /// <param name="value">The <see cref="bool"/> value.</param>
        public BooleanValue(bool value):base(value) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BooleanValue"/> class by deep copying
        /// the supplied <see cref="BooleanValue"/> class.
        /// </summary>
        /// <param name="source">The source <see cref="BooleanValue"/> class.</param>
        public BooleanValue(BooleanValue source) : base(source) { }


        /// <summary>
        /// Initializes a new instance of the <see cref="BooleanValue"/> class by implicitly
        /// converting the supplied <see cref="bool"/> value.
        /// </summary>
        /// <param name="value">The <see cref="bool"/> value.</param>
        public static implicit operator BooleanValue(bool value)
        {
            return new BooleanValue(value);
        }

    }
}
