using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Berry.Docx.Documents;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// Represent a style collection.
    /// </summary>
    public class StyleCollection : IEnumerable<Style>
    {
        #region Private Members
        private IEnumerable<Style> _styles;
        #endregion

        #region Constructors
        internal StyleCollection(IEnumerable<Style> styles)
        {
            _styles = styles;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the style at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        /// <returns>The style at the specified index in the current collection.</returns>
        public Style this[int index] => _styles.ElementAt(index);

        /// <summary>
        /// Gets the number of styles in the collection.
        /// </summary>
        public int Count => _styles.Count();
        #endregion

        #region Public Methods
        /// <summary>
        /// Searchs for the style with the specified stylename and type within the entire collection.
        /// </summary>
        /// <param name="name">The name of style.</param>
        /// <param name="type">The StyleType of style.</param>
        /// <returns>The style with the specified stylename and type</returns>
        public Style FindByName(string name, StyleType type)
        {
            return _styles.Where(s => s.Name.ToLower() == name.ToLower() && s.Type == type).FirstOrDefault();
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator that can be used to iterate through the collection.</returns>
        public IEnumerator<Style> GetEnumerator()
        {
            return _styles.GetEnumerator();
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _styles.GetEnumerator();
        }
        #endregion
    }
}
