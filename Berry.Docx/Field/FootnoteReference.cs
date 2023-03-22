using System;
using System.Collections.Generic;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// TODO
    /// </summary>
    public class FootnoteReference : ParagraphItem
    {
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.FootnoteReference _fnRef;
        internal FootnoteReference(Document doc, W.Run ownerRun, W.FootnoteReference fnRef)
            :base(doc, ownerRun, fnRef)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _fnRef = fnRef;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.FootnoteReference;

        public int Id
        {
            get
            {
                if (_fnRef.Id != null) return (int)_fnRef.Id;
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
            W.FootnoteReference fnRef = (W.FootnoteReference)_fnRef.CloneNode(true);
            run.RunProperties = _ownerRun.RunProperties?.CloneNode(true) as W.RunProperties; // copy format
            run.AppendChild(fnRef);
            return new FootnoteReference(_doc, run, fnRef);
        }
        #endregion
    }
}
