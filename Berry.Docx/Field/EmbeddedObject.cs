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
        private readonly W.EmbeddedObject _object;
        private readonly Ovml.OleObject _oleObject;
        internal EmbeddedObject(Document doc, W.Run ownerRun, W.EmbeddedObject obj)
            :base(doc, ownerRun, obj)
        {
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
    }
}
