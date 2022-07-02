using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Collections;

namespace Berry.Docx.Formatting
{
    public class ListStyle
    {
        private readonly Document _doc;
        private readonly W.AbstractNum _abstractNum;
        public ListStyle(Document doc, W.AbstractNum abstractNum)
        {
            _doc = doc;
            _abstractNum = abstractNum;
        }

        public string Name { get; set; }

        public ListLevelCollection Levels => new ListLevelCollection(GetLevels());

        private IEnumerable<ListLevel> GetLevels()
        {
            foreach (W.Level level in _abstractNum.Elements<W.Level>())
            {
                yield return new ListLevel(_doc, _abstractNum, level);
            }
        }
    }
}
