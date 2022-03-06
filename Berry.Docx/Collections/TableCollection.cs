using System.Collections.Generic;
using O = DocumentFormat.OpenXml;
using Berry.Docx.Documents;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// Represent a table collection.
    /// </summary>
    public class TableCollection : DocumentItemCollection
    {
        #region Constructors
        internal TableCollection(O.OpenXmlElement owner, IEnumerable<Table> tables) : base(owner, tables)
        {
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the table at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        /// <returns>The table at the specified index.</returns>
        public new Table this[int index] => (Table)base[index];
        #endregion

        #region Public Methods
        /// <summary>
        /// Returns the first table of the current collection.
        /// </summary>
        /// <returns>The first table in the current collection.</returns>
        public new Table First()
        {
            return (Table)base.First();
        }

        /// <summary>
        /// Returns the last table of the current collection.
        /// </summary>
        /// <returns>The last table in the current collection.</returns>
        public new Table Last()
        {
            return (Table)base.Last();
        }

        /// <summary>
        /// Determines whether this collection contains a specified table.
        /// </summary>
        /// <param name="table">The specified DocumentObject.</param>
        /// <returns>true if the collection contains the specified DocumentObject; otherwise, false.</returns>
        public bool Contains(Table table)
        {
            return base.Contains(table);
        }

        /// <summary>
        /// Adds the specified table to the end of the current collection.
        /// </summary>
        /// <param name="table">The table instance that was added.</param>
        public void Add(Table table)
        {
            base.Add(table);
        }

        /// <summary>
        /// Searchs for the specified table and returns the zero-based index of the first occurrence within the entire collection.
        /// </summary>
        /// <param name="table">The specified table.</param>
        /// <returns>The zero-based index of the first occurrence of table within the entire collection,if found; otherwise, -1.</returns>
        public int IndexOf(Table table)
        {
            return base.IndexOf(table);
        }

        /// <summary>
        /// Insert the specified table immediately to the specified index of the current collection.
        /// </summary>
        /// <param name="table">The inserted table instance.</param>
        /// <param name="index">The zero-based index.</param>
        public void InsertAt(Table table, int index)
        {
            base.InsertAt(table, index);
        }

        /// <summary>
        /// Removes the specified table immediately from the current collection.
        /// </summary>
        /// <param name="table"> The table instance that was removed.</param>
        public void Remove(Table table)
        {
            base.Remove(table);
        }
        #endregion
    }
}
