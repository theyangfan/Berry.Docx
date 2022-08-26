using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Represents the paragraph indentation.
    /// </summary>
    public class Indentation
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="Indentation"/> class.
        /// </summary>
        public Indentation() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="Indentation"/> class with specified parameters.
        /// </summary>
        /// <param name="val">The indentation value.</param>
        /// <param name="unit">The indentation measurement unit.</param>
        public Indentation(float val, IndentationUnit unit)
        {
            Val = val;
            Unit = unit;
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
        #endregion

        #region Public Methods
        public override string ToString()
        {
            return $"Indentation[{Val} {Unit}]";
        }
        #endregion

    }
}
