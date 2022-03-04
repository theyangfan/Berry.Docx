using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Represent the <see cref="bool"/> value.
    /// </summary>
    public class BooleanValue
    {
        private bool _value = false;
        /// <summary>
        /// Initializes a new instance of the <see cref="BooleanValue"/> class.
        /// </summary>
        public BooleanValue() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="BooleanValue"/> class using the supplied <see cref="bool"/> value.
        /// </summary>
        /// <param name="value">The <see cref="bool"/> value.</param>
        public BooleanValue(bool value)
        {
            _value = value;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="BooleanValue"/> class by deep copying
        /// the supplied <see cref="BooleanValue"/> class.
        /// </summary>
        /// <param name="source">The source <see cref="BooleanValue"/> class.</param>
        public BooleanValue(BooleanValue source)
        {
            _value = source.Val;
        }

        /// <summary>
        /// Gets or sets the inner <see cref="bool"/> value.
        /// </summary>
        public bool Val
        {
            get => _value;
            set => _value = value;
        }

        /// <summary>
        /// Implicitly converting the <see cref="BooleanValue"/> value to <see cref="bool"/> value.
        /// </summary>
        /// <param name="value">The <see cref="BooleanValue"/> value.</param>
        public static implicit operator bool(BooleanValue value)
        {
            return value.Val;
        }

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
