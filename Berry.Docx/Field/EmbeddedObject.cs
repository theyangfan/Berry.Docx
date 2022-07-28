using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;

namespace Berry.Docx.Field
{
    public class EmbeddedObject : ParagraphItem
    {
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.EmbeddedObject _object;
        private readonly Ovml.OleObject _oleObject;
        internal EmbeddedObject(Document doc, W.Run ownerRun, W.EmbeddedObject obj)
            :base(doc, ownerRun, obj)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _object = obj;
            _oleObject = obj.Elements<Ovml.OleObject>().FirstOrDefault();
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.EmbeddedObject;

        public OleObjectType OleType
        {
            get
            {
                if (_oleObject?.Type != null)
                    return _oleObject.Type.Value == Ovml.OleValues.Embed ? OleObjectType.Embed : OleObjectType.Link;
                return OleObjectType.Embed;
            }
        }

        public string OleProgId => _oleObject?.ProgId ?? string.Empty;

        #region Public Methods
        /// <summary>
        /// Creates a duplicate of the object.
        /// </summary>
        /// <returns>The cloned object.</returns>
        public override DocumentObject Clone()
        {
            W.Run run = new W.Run();
            W.EmbeddedObject embobj = (W.EmbeddedObject)_object.CloneNode(true);
            run.RunProperties = _ownerRun.RunProperties?.CloneNode(true) as W.RunProperties; // copy format
            run.AppendChild(embobj);
            return new EmbeddedObject(_doc, run, embobj);
        }
        #endregion
    }
}
