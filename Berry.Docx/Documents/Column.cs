using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    public class Column
    {
        private readonly W.Column _column;
        private float _width;
        private float _space;
        internal Column(Document doc, W.Column column)
        {
            _column = column;
        }
        public Column(Document doc)
        {
            _column = new W.Column() { Width = "0", Space = "0" };
        }

        internal W.Column XElement => _column;
        public float Width
        {
            get
            {
                if (_column.Width == null) return 0;
                return (_column.Width.ToString().ToInt() / 20.0F).Round(2);
            }
            set
            {
                _column.Width = (value * 20).Round(0).ToString();
            }
        }
        public float Spacing
        {
            get
            {
                if (_column.Space == null) return 0;
                return (_column.Space.ToString().ToInt() / 20.0F).Round(2);
            }
            set
            {
                _column.Space = (value * 20).Round(0).ToString();
            }
        }
    }
}
