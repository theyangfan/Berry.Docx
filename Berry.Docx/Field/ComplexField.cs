using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    /// <summary>
    /// TODO
    /// </summary>
    internal class ComplexField : ParagraphItem
    {
        private readonly Document _doc;
        private readonly W.Run _begin;
        private readonly W.Run _end;
        internal ComplexField(Document doc, W.Run begin, W.Run end) : base(doc, begin)
        {
            _doc = doc;
            _begin = begin;
            _end = end;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.ComplexField;

        public string Code
        {
            get;
        }

        public string Result
        {
            get;
        }
    }
}
