using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Represents the paragraph Margins.
    /// </summary>
    public class MarginsF
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="MarginsF"/> class.
        /// </summary>
        public MarginsF() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="MarginsF"/> class with specified parameters.
        /// </summary>
        /// <param name="left">The left margin.</param>
        /// <param name="right">The right margin.</param>
        /// <param name="top">The top margin.</param>
        /// <param name="bottom">The bottom margin.</param>
        public MarginsF(float left, float right, float top, float bottom)
        {
            Left = left;
            Right = right;
            Top = top;
            Bottom = bottom;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The left margin.
        /// </summary>
        public float Left { get; set; }

        /// <summary>
        /// The right margin.
        /// </summary>
        public float Right { get; set; }

        /// <summary>
        /// The top margin.
        /// </summary>
        public float Top { get; set; }

        /// <summary>
        /// The bottom margin.
        /// </summary>
        public float Bottom { get; set; }
        #endregion

        #region Public Methods
        public override string ToString()
        {
            return "{" + $"Left:{Left}, Right:{Right}, Top:{Top}, Bottom:{Bottom}" + "}";
        }
        #endregion

    }
}
