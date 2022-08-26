using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Represents the paragraph special indentation.
    /// </summary>
    public class SpecialIndentation
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="SpecialIndentation"/> class.
        /// </summary>
        public SpecialIndentation() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="SpecialIndentation"/> class with specified parameters.
        /// </summary>
        /// <param name="type">The indentation type.</param>
        /// <param name="val">The indentation value.</param>
        /// <param name="unit">The indentation measurement unit.</param>
        public SpecialIndentation(SpecialIndentationType type, float val, IndentationUnit unit)
        {
            Val = val;
            Unit = unit;
            Type = type;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets or sets the indentation value.
        /// </summary>
        public float Val { get; set; }

        /// <summary>
        /// Gets or sets the indentation measurement unit.
        /// </summary>
        public IndentationUnit Unit { get; set; }

        /// <summary>
        /// Gets or sets the indentation type.
        /// </summary>
        public SpecialIndentationType Type { get; set; }
        #endregion

        #region Public Methods
        public override string ToString()
        {
            if (Type != SpecialIndentationType.None)
                return $"SpecialIndentation[{Type}: {Val} {Unit}]";
            else
                return $"SpecialIndentation[{Type}]";
        }
        #endregion

    }
}
