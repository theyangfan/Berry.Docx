using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// Represents a hyperlink.
    /// </summary>
    public class Hyperlink : ParagraphItem
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Hyperlink _hyperlink;
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new Hyperlink instance with the specified type and target.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="type">The type of the target.</param>
        /// <param name="target">The external hyperlink address or bookmark name.</param>
        public Hyperlink(Document doc, HyperlinkTargetType type, string target)
            : this(doc, ParagraphItemGenerator.GenerateHyperlink(doc, type, target))
        {
        }

        internal Hyperlink(Document doc, W.Hyperlink hyperlink) : base(doc, hyperlink)
        {
            _doc = doc;
            _hyperlink = hyperlink;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Returns the type of the object.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Hyperlink;

        /// <summary>
        /// Returns the type of the target.
        /// </summary>
        public HyperlinkTargetType TargetType
        {
            get
            {
                if(_hyperlink.Id != null)
                {
                    return HyperlinkTargetType.ExternalAddress;
                }
                else if(_hyperlink.Anchor != null)
                {
                    return HyperlinkTargetType.Bookmark;
                }
                return HyperlinkTargetType.Default;
            }
        }

        /// <summary>
        /// Gets or sets the target value. If the TargetType is ExternalAddress, 
        /// the value is the external hyperlink address. If the TargetType is 
        /// Bookmark, the value is the name of the bookmark. Otherwise, returns
        /// string.Empty.
        /// </summary>
        public string Target
        {
            get
            {
                if(_hyperlink.Id != null)
                {
                    var relationship = _doc.Package.MainDocumentPart.HyperlinkRelationships
                        .Where(r => r.Id == _hyperlink.Id).FirstOrDefault();
                    if (relationship != null) return relationship.Uri.OriginalString;
                }
                else if(_hyperlink.Anchor != null)
                {
                    return _hyperlink.Anchor;
                }
                return string.Empty;
            }
            set
            {
                if(_hyperlink.Id != null)
                {
                    _doc.Package.MainDocumentPart.DeleteReferenceRelationship(_hyperlink.Id);
                    var newRelationship = _doc.Package.MainDocumentPart.AddHyperlinkRelationship(new Uri(value), true);
                    _hyperlink.Id = newRelationship.Id;
                }
                else if(_hyperlink.Anchor != null)
                {
                    _hyperlink.Anchor = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicates whether the hyperlink target
        /// shall be added to a list of viewed hyperlinks when it is invoked. 
        /// </summary>
        public bool AddToViewedHistory
        {
            get
            {
                if(_hyperlink.History == null) return false;
                return _hyperlink.History.Value;
            }
            set
            {
                if (value) _hyperlink.History = true;
                else _hyperlink.History = null;
            }
        }

        /// <summary>
        /// Gets or sets the text to display for the current hyperlink.
        /// </summary>
        public string Text
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                foreach(var tr in ChildObjects.OfType<TextRange>())
                {
                    sb.Append(tr.Text);
                }
                return sb.ToString();
            }
            set
            {
                ChildObjects.Clear();
                TextRange tr= new TextRange(_doc, value);
                ChildObjects.Add(tr);
            }
        }
        #endregion
    }
}
