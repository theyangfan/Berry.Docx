using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the numbering format.
    /// </summary>
    public class ListFormat
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.AbstractNum _abstractNum;
        private readonly W.Level _curLevel;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the NumberingFormat class using the supplied OpenXML AbstractNum element.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="num"></param>
        /// 
        internal ListFormat(Document doc, W.Paragraph ownerParagraph)
        {

        }

        internal ListFormat(Document doc, W.Style ownerStyle)
        {

        }

        internal ListFormat(Document doc, W.AbstractNum num, int levelIndex)
        {
            _doc = doc;
            _abstractNum = num;
            _curLevel = num.Elements<W.Level>().Where(l => l.LevelIndex == levelIndex).FirstOrDefault();
        }
        internal ListFormat(Document doc, W.AbstractNum num, string styleId)
        {
            _doc = doc;
            _abstractNum = num;
            _curLevel = num.Elements<W.Level>().Where(l => l.ParagraphStyleIdInLevel?.Val == styleId).FirstOrDefault();
        }
        internal ListFormat(Document doc, ListFormat format, string styleId)
        {
            _doc = doc;
            _abstractNum = format.AbstractNum;
            _curLevel = _abstractNum.Elements<W.Level>().Where(l => l.ParagraphStyleIdInLevel?.Val == styleId).FirstOrDefault();
        }
        #endregion

        #region Public Properties

        public int ListLevelNumber
        {
            get => _curLevel.LevelIndex.Value + 1;
        }

        public ListStyle CurrentStyle => new ListStyle(_doc, _abstractNum);

        public ListLevel CurrentLevel => new ListLevel(_doc, _abstractNum, _curLevel);

        public void ApplyStyle(ListStyle style)
        {

        }

        #endregion

        #region Internal Properties
        internal W.AbstractNum AbstractNum => _abstractNum;
        #endregion

    }
}
