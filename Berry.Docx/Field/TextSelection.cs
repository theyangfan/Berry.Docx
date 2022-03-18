using System;
using System.Collections.Generic;
using System.Text;
using Berry.Docx.Documents;

namespace Berry.Docx.Field
{
    public class TextSelection
    {
        private readonly Paragraph _ownerParagraph;
        private readonly int _startCharPos;
        private readonly int _endCharPos;
        public TextSelection(Paragraph ownerParagraph, int startCharPos, int endCharPos)
        {
            _ownerParagraph = ownerParagraph;
            _startCharPos = startCharPos;
            _endCharPos = endCharPos;
        }

        public string Text => _ownerParagraph.Text.Substring(_startCharPos, _endCharPos - _startCharPos + 1);

        public TextRange GetAsOneRange()
        {
            for(int i = 0; i < _ownerParagraph.ChildObjects.Count; i++)
            {

            }
            return null;
        }
    }
}
