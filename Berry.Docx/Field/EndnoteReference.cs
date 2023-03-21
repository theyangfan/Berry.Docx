using System;
using System.Collections.Generic;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// TODO
    /// </summary>
    public class EndnoteReference : ParagraphItem
    {
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.EndnoteReference _enRef;
        internal EndnoteReference(Document doc, W.Run ownerRun, W.EndnoteReference enRef)
            :base(doc, ownerRun, enRef)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _enRef = enRef;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.EndnoteReference;

        public int Id
        {
            get
            {
                if (_enRef.Id != null) return (int)_enRef.Id;
                return -1;
            }
        }

        #region Public Methods
        /// <summary>
        /// Creates a duplicate of the object.
        /// </summary>
        /// <returns>The cloned object.</returns>
        public override DocumentObject Clone()
        {
            W.Run run = new W.Run();
            W.EndnoteReference enRef = (W.EndnoteReference)_enRef.CloneNode(true);
            run.RunProperties = _ownerRun.RunProperties?.CloneNode(true) as W.RunProperties; // copy format
            run.AppendChild(enRef);
            return new EndnoteReference(_doc, run, enRef);
        }
        #endregion
    }
}
