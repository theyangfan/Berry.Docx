using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// Represent a DocumentObject collection.
    /// </summary>
    public abstract class DocumentObjectCollection : IEnumerable
    {
        #region Private Members
        private IEnumerable<DocumentObject> _objects;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the DocumentObjectCollection class using the supplied collection.
        /// </summary>
        /// <param name="objects">The DocumentObject collection</param>
        public DocumentObjectCollection(IEnumerable<DocumentObject> objects)
        {
            _objects = objects;
        }
        #endregion

        #region Public Properties

        /// <summary>
        /// Gets the DocumentObject at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        /// <returns>The DocumentObject at the specified index.</returns>
        public DocumentObject this[int index] => _objects.ElementAt(index);

        /// <summary>
        /// Gets the number of DocumentObjects in the collection.
        /// </summary>
        public virtual int Count => _objects.Count();
        #endregion

        #region Public Methods
        /// <summary>
        /// Determines whether this collection contains a specified DocumentObject.
        /// </summary>
        /// <param name="obj">The specified DocumentObject.</param>
        /// <returns>true if the collection contains the specified DocumentObject; otherwise, false.</returns>
        public virtual bool Contains(DocumentObject obj)
        {
            return _objects.Contains(obj);
        }

        /// <summary>
        /// Searchs for the specified object and returns the zero-based index of the first occurrence within the entire collection.
        /// </summary> 
        /// <param name="obj">The specified DocumentObject.</param>
        /// <returns>The zero-based index of the first occurrence of item within the entire collection,if found; otherwise, -1.</returns>
        public virtual int IndexOf(DocumentObject obj)
        {
            return _objects.ToList().IndexOf(obj);
        }

        /// <summary>
        /// Filters the elements of an <see cref="IEnumerable"/> based on a specified
        /// type.
        /// </summary>
        /// <typeparam name="T">The type to filter the elements of the sequence on.</typeparam>
        /// <returns>
        /// An <see cref="IEnumerable"/> that contains elements from the input
        /// sequence of type TResult.
        /// </returns>
        public IEnumerable<T> OfType<T>()
        {
            return _objects.OfType<T>();
        }

        /// <summary>
        /// Adds the specified object to the end of the current collection.
        /// </summary>
        /// <param name="obj">The DocumentObject instance that was added.</param>
        public abstract void Add(DocumentObject obj);

        /// <summary>
        /// Insert the specified object immediately to the specified index of the current collection.
        /// </summary>
        /// <param name="obj">The inserted DocumentObject instance.</param>
        /// <param name="index">The zero-based index.</param>
        public abstract void InsertAt(DocumentObject obj, int index);

        /// <summary>
        /// Removes the specified DocumentObject immediately from the current collection.
        /// </summary>
        /// <param name="obj"> The DocumentObject instance that was removed. </param>
        public abstract void Remove(DocumentObject obj);

        /// <summary>
        /// Removes the DocumentObject at the zero-based index immediately from the current collection.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        public abstract void RemoveAt(int index);

        /// <summary>
        /// Removes all items of the current collection.
        /// </summary>
        public abstract void Clear();

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator that can be used to iterate through the collection.</returns>
        public IEnumerator GetEnumerator()
        {
            return _objects.GetEnumerator();
        }
        #endregion

    }
}
