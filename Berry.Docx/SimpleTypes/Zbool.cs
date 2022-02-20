using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Represent the <see cref="bool"/> value.
    /// </summary>
    public class ZBool
    {
        private bool _value = false;
        /// <summary>
        /// Initializes a new instance of the <see cref="ZBool"/> class.
        /// </summary>
        public ZBool() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ZBool"/> class using the supplied <see cref="bool"/> value.
        /// </summary>
        /// <param name="value">The <see cref="bool"/> value.</param>
        public ZBool(bool value)
        {
            _value = value;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ZBool"/> class by deep copying
        /// the supplied <see cref="ZBool"/> class.
        /// </summary>
        /// <param name="source">The source <see cref="ZBool"/> class.</param>
        public ZBool(ZBool source)
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
        /// Implicitly converting the <see cref="ZBool"/> value to <see cref="bool"/> value.
        /// </summary>
        /// <param name="value">The <see cref="ZBool"/> value.</param>
        public static implicit operator bool(ZBool value)
        {
            return value.Val;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ZBool"/> class by implicitly
        /// converting the supplied <see cref="bool"/> value.
        /// </summary>
        /// <param name="value">The <see cref="bool"/> value.</param>
        public static implicit operator ZBool(bool value)
        {
            return new ZBool(value);
        }
    }
}
