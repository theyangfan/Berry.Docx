using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

namespace Berry.Docx
{
    /// <summary>
    /// Represent a <see cref="Color"/> value;
    /// </summary>
    public class ColorValue
    {
        #region Private Members
        private Color _color = Color.Empty;
        private bool _auto = false;
        #endregion

        #region Static Members
        public static ColorValue Auto = new ColorValue() { IsAuto = true };
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a <see cref="Color.Empty"/> value.
        /// </summary>
        public ColorValue() { }

        /// <summary>
        /// Creates an instance from the RRGGBB hex string.
        /// </summary>
        /// <param name="rgb">The RRGGBB hex string.</param>
        public ColorValue(string rgb)
        {
            _color = ColorConverter.FromHex(rgb);
            _auto = rgb == "auto";
        }

        /// <summary>
        /// Creates an instance with the specified <see cref="Color"/>
        /// </summary>
        /// <param name="color"></param>
        public ColorValue(Color color)
        {
            _color = color;
        }

        /// <summary>
        /// Creates an instance with the specified <see cref="ColorValue"/>
        /// </summary>
        /// <param name="source"></param>
        public ColorValue(ColorValue source)
        {
            _color = source.Val;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets or sets a value indicating whether the color is auto.
        /// </summary>
        public bool IsAuto
        {
            get => _auto;
            set => _auto = value;
        }

        /// <summary>
        /// Gets or sets the <see cref="Color"/> value.
        /// </summary>
        public Color Val
        {
            get => _color;
            set => _color = value;
        }
        #endregion

        #region Public Methods

        public static implicit operator Color(ColorValue value)
        {
            return value.Val;
        }

        public static implicit operator ColorValue(Color value)
        {
            return new ColorValue(value);
        }

        public static implicit operator ColorValue(string rgb)
        {
            return new ColorValue(rgb);
        }

        public override string ToString()
        {
            return _auto ? "auto" : ColorConverter.ToHex(_color);
        }
        #endregion
    }
}
