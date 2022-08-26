using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Represents the paragraph  spacing.
    /// </summary>
    public class Spacing
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="Spacing"/> class.
        /// </summary>
        public Spacing() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="Spacing"/> class with specified parameters.
        /// </summary>
        /// <param name="val">The spacing value.</param>
        /// <param name="unit">The spacing measurement unit.</param>
        public Spacing(float val, SpacingUnit unit)
        {
            Val = val;
            Unit = unit;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets or sets the spacing value.
        /// </summary>
        public float Val { get; set; }

        /// <summary>
        /// Gets or sets the spacing measurement unit.
        /// </summary>
        public SpacingUnit Unit { get; set; }
        #endregion

        #region Public Methods
        public override string ToString()
        {
            return $"Spacing[{Val} {Unit}]";
        }
        #endregion

    }
}
