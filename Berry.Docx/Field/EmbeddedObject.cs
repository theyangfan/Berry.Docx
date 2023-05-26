using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using P = DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
#if NET35_OR_GREATER
using System.Drawing;
#elif NETCOREAPP2_1_OR_GREATER || NETSTANDARD2_0_OR_GREATER
using SixLabors.ImageSharp;
#endif

namespace Berry.Docx.Field
{
    /// <summary>
    /// This class represents an embedded object in the paragraph.
    /// </summary>
    public class EmbeddedObject : ParagraphItem
    {
#region Private Members
        private readonly Document _doc;
        private readonly W.Run _ownerRun;
        private readonly W.EmbeddedObject _object;
        private readonly V.Shape _shape;
        private readonly Ovml.OleObject _oleObject;
        private readonly ShapeStyles _shapeStyles;
#endregion

#region Constructors
        internal EmbeddedObject(Document doc, W.Run ownerRun, W.EmbeddedObject obj)
            :base(doc, ownerRun, obj)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            _object = obj;
            _oleObject = obj.Elements<Ovml.OleObject>().FirstOrDefault();
            _shape = obj.GetFirstChild<V.Shape>();
            _shapeStyles = new ShapeStyles(_shape);
        }
#endregion

#region Public Properties
        /// <summary>
        /// Gets the type of the current object.
        /// </summary>
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.EmbeddedObject;

        /// <summary>
        /// 
        /// </summary>
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

        public float Width
        {
            get
            {
                if (_shapeStyles["width"].EndsWith("pt"))
                {
                    string width = _shapeStyles["width"].Replace("pt", "");
                    return Convert.ToSingle(width);
                }
                else if (_shapeStyles["width"].EndsWith("in"))
                {
                    string width = _shapeStyles["width"].Replace("in", "");
                    return Convert.ToSingle(width) * 72;
                }
                return 0;
            }
            set
            {
                using (var s = GetStream())
                {
#if NET35_OR_GREATER
                    var img = Image.FromStream(s);
#elif NETCOREAPP2_1_OR_GREATER || NETSTANDARD2_0_OR_GREATER
                    var img = Image.Load(s);
#endif
                    _object.DxaOriginal = Convert.ToInt32(img.Width / 96.0f * 72 * 20).ToString();
                    img.Dispose();
                }
                _shapeStyles["width"] = $"{Math.Round(value, 2)}pt";
            }
        }

        
        public float Height
        {
            get
            {
                if (_shapeStyles.Contains("height"))
                {
                    if (_shapeStyles["height"].EndsWith("pt"))
                    {
                        string height = _shapeStyles["height"].Replace("pt", "");
                        return Convert.ToSingle(height);
                    }
                    else if (_shapeStyles["height"].EndsWith("in"))
                    {
                        string height = _shapeStyles["height"].Replace("in", "");
                        return Convert.ToSingle(height) * 72;
                    }
                }
                return 0;
            }
            set
            {
                using(var s = GetStream())
                {
#if NET35_OR_GREATER
                    var img = Image.FromStream(s);
#elif NETCOREAPP2_1_OR_GREATER || NETSTANDARD2_0_OR_GREATER
                    var img = Image.Load(s);
#endif
                    _object.DyaOriginal = Convert.ToInt32(img.Height / 96.0f * 72 * 20).ToString();
                    img.Dispose();
                }
                _shapeStyles["height"] = $"{Math.Round(value, 2)}pt";
            }
        }

#endregion

#region Public Methods
        public Stream GetStream()
        {
            string rId = _shape.GetFirstChild<V.ImageData>().RelationshipId.Value;
            if (string.IsNullOrEmpty(rId)) return null;
            P.ImagePart imagePart = (P.ImagePart)_doc.Package.MainDocumentPart.GetPartById(rId);
            if (imagePart == null) return null;
            return imagePart.GetStream();
        }

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

    internal class ShapeStyles : IEnumerable<KeyValuePair<string, string>>
    {
        private readonly V.Shape _shape;
        private Dictionary<string, string> _styles;
        internal ShapeStyles(V.Shape shape)
        {
            _shape = shape;
            _styles = new Dictionary<string, string>();
            if (!string.IsNullOrEmpty(shape?.Style?.Value))
            {
                foreach (var style in shape.Style.Value.Split(';'))
                {
                    string key = style.Split(':')?.FirstOrDefault();
                    string value = style.Split(':')?.LastOrDefault();
                    if (string.IsNullOrEmpty(key)) continue;
                    _styles[key] = value;
                }
            }
        }

        public string this[string key]
        {
            get
            {
                return _styles[key];
            }
            set
            {
                _styles[key] = value;
                StringBuilder styleStr = new StringBuilder();
                foreach(var style in _styles)
                {
                    styleStr.Append($"{style.Key}:{style.Value};");
                }
                _shape.Style = styleStr.ToString();
            }
        }

        public bool Contains(string key)
        {
            return _styles.ContainsKey(key);
        }

        public IEnumerator<KeyValuePair<string, string>> GetEnumerator()
        {
            return _styles.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _styles.GetEnumerator();
        }
    }
}
