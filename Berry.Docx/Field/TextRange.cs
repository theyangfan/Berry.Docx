using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;
using Berry.Docx.Collections;
using Berry.Docx.Utils;
using Berry.Docx.Formatting;

namespace Berry.Docx.Field
{
    public class TextRange : DocumentElement
    {
        private Document _doc;
        private W.Run _run;
        private W.Text _text;
        private CharacterFormat _cFormat;

        public TextRange(Document doc):this(doc, RunGenerator.Generate(""))
        {}
        internal TextRange(Document doc, W.Run run) : base(doc, run)
        {
            _doc = doc;
            _run = run;
            _text = run.Elements<W.Text>().FirstOrDefault();
            _cFormat = new CharacterFormat(doc, run);
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.TextRange;

        public string Text
        {
            get
            {
                if (_text != null)
                    return _text.Text;
                return string.Empty;
            }
            set
            {
                if(_text == null)
                {
                    _text = new W.Text();
                    _run.AddChild(_text);
                }
                _text.Text = value;
            }
        }

        public CharacterFormat CharacterFormat => _cFormat;

    }
}
