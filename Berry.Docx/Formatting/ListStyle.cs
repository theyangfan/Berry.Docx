using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
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

        internal W.NumberingInstance NumberingInstance
        {
            get
            {
                W.Numbering numbering = _doc.Package.MainDocumentPart.NumberingDefinitionsPart?.Numbering;
                if (numbering != null)
                {
                    W.NumberingInstance num = numbering.Elements<W.NumberingInstance>()
                        .Where(n => n.AbstractNumId.Val == _abstractNum.AbstractNumberId).FirstOrDefault();
                    if (num != null) return num;
                }
                W.NumberingInstance numberingInstance = new W.NumberingInstance()
                {
                    NumberID = IDGenerator.GenerateNumId(_doc)
                };
                numberingInstance.AbstractNumId = new W.AbstractNumId() { Val = _abstractNum.AbstractNumberId };
                return numberingInstance;
            }
        }
        private IEnumerable<ListLevel> GetLevels()
        {
            foreach (W.Level level in _abstractNum.Elements<W.Level>())
            {
                yield return new ListLevel(_doc, _abstractNum, level);
            }
        }
    }
}
