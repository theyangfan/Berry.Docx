using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Formatting;

namespace Berry.Docx.Collections
{
    public class ListStyleCollection : IEnumerable<ListStyle>
    {
        #region Private Members
        private readonly Document _doc;
        private IEnumerable<ListStyle> _styles;
        private static Dictionary<string, int> _listStyleNames = new Dictionary<string, int>();
        #endregion

        #region Constructors
        internal ListStyleCollection(Document doc)
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
        public ListStyle this[int index] => _styles.ElementAt(index);

        /// <summary>
        /// Gets the number of styles in the collection.
        /// </summary>
        public int Count => _styles.Count();
        #endregion

        public void Add(ListStyle style)
        {
            if (GetStyles().Where(s => s.NumberID == style.NumberID || s.AbstractNumberID == style.AbstractNumberID).Any()) return;

            if (_doc.Package.MainDocumentPart.NumberingDefinitionsPart == null)
            {
                PartGenerator.AddNumberingPart(_doc, IDGenerator.GenerateRelationshipID(_doc));
            }
            W.Numbering numbering = _doc.Package.MainDocumentPart.NumberingDefinitionsPart.Numbering;
            if (numbering.Elements<W.AbstractNum>().Any())
            {
                numbering.Elements<W.AbstractNum>().Last().InsertAfterSelf(style.AbstractNum);
            }
            else
            {
                numbering.Append(style.AbstractNum);
            }
            numbering.Append(style.NumberingInstance);
            if (!string.IsNullOrWhiteSpace(style.Name))
                _listStyleNames[style.Name] = style.AbstractNumberID;
        }

        public ListStyle FindByName(string styleName)
        {
            return _styles.Where(s => s.Name == styleName).FirstOrDefault();
        }

        public IEnumerator<ListStyle> GetEnumerator()
        {
            return _styles.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        private IEnumerable<ListStyle> GetStyles()
        {
            if (_doc.Package.MainDocumentPart.NumberingDefinitionsPart?.Numbering == null) yield break;
            foreach (W.AbstractNum num in _doc.Package.MainDocumentPart.NumberingDefinitionsPart.Numbering.Elements<W.AbstractNum>())
            {
                ListStyle style = new ListStyle(_doc, num);
                if (_listStyleNames.ContainsValue(style.AbstractNumberID))
                {
                    style.Name = _listStyleNames.Where(p => p.Value == style.AbstractNumberID).First().Key;
                }
                yield return style;
            }
        }
    }
}
