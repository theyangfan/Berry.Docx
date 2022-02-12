using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// Represent a section collection.
    /// </summary>
    public class SectionCollection : IEnumerable
    {
        #region Private Members
        private IEnumerable<Section> _sections;
        #endregion

        #region Constructors
        internal SectionCollection(IEnumerable<Section> sections)
        {
            _sections = sections;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the section at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        /// <returns>The section at the specified index in the current collection.</returns>
        public Section this[int index] => _sections.ElementAt(index);

        /// <summary>
        /// Gets the number of sections in the collection.
        /// </summary>
        public int Count => _sections.Count();
        #endregion

        #region Public Methods
        /// <summary>
        /// Searchs for the specified section and returns the zero-based index of the first occurrence within the entire collection.
        /// </summary> 
        /// <param name="section">The specified section.</param>
        /// <returns>The zero-based index of the first occurrence of section within the entire collection,if found; otherwise, -1.</returns>
        public int IndexOf(Section section) => _sections.ToList().IndexOf(section);

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator that can be used to iterate through the collection.</returns>
        public IEnumerator GetEnumerator()
        {
            return _sections.GetEnumerator();
        }
        #endregion
    }
}
