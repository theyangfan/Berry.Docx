using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Berry.Docx.Formatting;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// Represent a style collection.
    /// </summary>
    public class StyleCollection : IEnumerable<Style>
    {
        #region Private Members
        private readonly Document _doc;
        private IEnumerable<Style> _styles;
        #endregion

        #region Constructors
        internal StyleCollection(Document doc)
        {
            _doc = doc;
            _styles = GetStyles();
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
        /// Adds the specified style to the document.
        /// <para>将指定的样式添加到文档中.</para>
        /// </summary>
        /// <param name="style">The specified style.</param>
        public void Add(Style style)
        {
            if(FindByName(style.Name, style.Type) == null)
            {
                _doc.Package.MainDocumentPart.StyleDefinitionsPart.Styles.Append(style.XElement);
            }
        }

        /// <summary>
        /// Searchs for the style with the specified stylename and type within the entire collection.
        /// </summary>
        /// <param name="name">The name of style.</param>
        /// <param name="type">The StyleType of style.</param>
        /// <returns>The style with the specified stylename and type</returns>
        public Style FindByName(string name, StyleType type)
        {
            name = Style.NameToBuiltInString(name);
            return _styles.Where(s => s.Name.ToLower() == name && s.Type == type).FirstOrDefault();
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

        #region Private Methods
        private IEnumerable<Style> GetStyles()
        {
            foreach (W.Style style in _doc.Package.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<W.Style>())
            {
                if (style.Type == W.StyleValues.Paragraph)
                    yield return new ParagraphStyle(_doc, style);
                else if (style.Type == W.StyleValues.Character)
                    yield return new CharacterStyle(_doc, style);
                else if (style.Type == W.StyleValues.Table)
                    yield return new TableStyle(_doc, style);
            }
        }
        #endregion
    }
}
