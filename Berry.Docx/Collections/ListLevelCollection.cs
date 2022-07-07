// Copyright (c) theyangfan. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using Berry.Docx.Formatting;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// Represent a collection of list levels.
    /// </summary>
    public class ListLevelCollection : IEnumerable<ListLevel>
    {
        #region Private Members
        private readonly IEnumerable<ListLevel> _levels;
        #endregion
        
        #region Constructors
        internal ListLevelCollection(IEnumerable<ListLevel> levels)
        {
            _levels = levels;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the level at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        /// <returns>The list level.</returns>
        public ListLevel this[int index] => _levels.ElementAt(index);

        /// <summary>
        /// Gets the levels count.
        /// </summary>
        public int Count => _levels.Count();
        #endregion

        #region Public Methods
        public IEnumerator<ListLevel> GetEnumerator()
        {
            return _levels.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        #endregion

    }
}
