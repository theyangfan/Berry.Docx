using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OOxml = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the numbering format.
    /// </summary>
    public class NumberingFormat
    {
        #region Private Members
        private OOxml.Level _lvl = null;
        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the NumberingFormat class using the supplied OpenXML Level element.
        /// </summary>
        /// <param name="lvl"></param>
        internal NumberingFormat(OOxml.Level lvl)
        {
            _lvl = lvl;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets start number.
        /// </summary>
        public int Start
        {
            get => _lvl.StartNumberingValue.Val;
        }

        /// <summary>
        /// Gets number style.
        /// </summary>
        public OOxml.NumberFormatValues Style
        {
            get => _lvl.NumberingFormat.Val;
        }

        /// <summary>
        /// Gets number format text.
        /// </summary>
        public string Format
        {
            get => _lvl.LevelText.Val;
        }
        #endregion
    }
}
