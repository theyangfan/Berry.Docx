using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    /// <summary>
    /// Represents the paragraph line spacing.
    /// </summary>
    public class LineSpacing
    {
        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="LineSpacing"/> class.
        /// </summary>
        public LineSpacing() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="LineSpacing"/> class with specified parameters.
        /// </summary>
        /// <param name="val">The line spacing value.</param>
        /// <param name="rule">The line spacing rule.</param>
        public LineSpacing(float val, LineSpacingRule rule)
        {
            Val = val;
            Rule = rule;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets or sets the line spacing value.
        /// </summary>
        public float Val { get; set; }

        /// <summary>
        /// Gets or sets the line spacing rule.
        /// </summary>
        public LineSpacingRule Rule { get; set; }
        #endregion

        #region Public Methods
        public override string ToString()
        {
            return $"LineSpacing[{Val} {Rule}]";
        }
        #endregion
    }
}
