using System;
using System.Collections.Generic;
using System.Linq;
using Berry.Docx.Documents;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    public class TextSelection
    {
        private readonly Paragraph _ownerParagraph;
        private readonly int _startCharPos;
        private readonly int _endCharPos;
        private List<TextRange> _ranges;
        public TextSelection(Paragraph ownerParagraph, int startCharPos, int endCharPos)
        {
            _ownerParagraph = ownerParagraph;
            _startCharPos = startCharPos;
            _endCharPos = endCharPos;
            _ranges = new List<TextRange>();
        }

        public string Text => _ownerParagraph.Text.Substring(_startCharPos, _endCharPos - _startCharPos + 1);

        public TextRange GetAsOneRange()
        {
            Console.WriteLine($"{_ownerParagraph.Text},{_startCharPos}->{_endCharPos}");
            string text = string.Empty;
            TextRange startRange = null;
            TextRange endRange = null;
            int startPos = -1; ;
            int endPos = -1;
            IEnumerable<TextRange> ranges = _ownerParagraph.ChildObjects.OfType<TextRange>();
            for (int i = 0; i < ranges.Count(); i++)
            {
                TextRange tr = ranges.ElementAt(i);
                Console.WriteLine(tr.Text);
                if(text.Length <= _startCharPos && (text+tr.Text).Length > _startCharPos)
                {
                    _ranges.Add(tr);
                    startPos = _startCharPos - text.Length;
                }
                if(text.Length <= _endCharPos && (text+tr.Text).Length > _endCharPos)
                {
                    if(!_ranges.Contains(tr))
                        _ranges.Add(tr);
                    endPos = _endCharPos - text.Length;
                }
                text += tr.Text;
            }
            Console.WriteLine("---------------");
            foreach (TextRange tr in _ranges)
            {
                Console.WriteLine(tr.Text);
            }
            Console.WriteLine(startPos);
            Console.WriteLine(endPos);
            if(startRange == endRange)
            {
                
            }
            else
            {

            }

            return null;
        }
    }
}
