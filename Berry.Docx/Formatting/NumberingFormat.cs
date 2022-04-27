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
    public class NumberingFormat
    {
        #region Private Members
        private readonly W.AbstractNum _abstractNum;
        private readonly W.Level _curLevel;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the NumberingFormat class using the supplied OpenXML AbstractNum element.
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="num"></param>
        internal NumberingFormat(Document doc, W.AbstractNum num, int levelIndex)
        {
            _abstractNum = num;
            _curLevel = num.Elements<W.Level>().Where(l => l.LevelIndex == levelIndex).FirstOrDefault();
        }
        internal NumberingFormat(Document doc, W.AbstractNum num, string styleId)
        {
            _abstractNum = num;
            _curLevel = num.Elements<W.Level>().Where(l => l.ParagraphStyleIdInLevel?.Val == styleId).FirstOrDefault();
        }
        internal NumberingFormat(Document doc, NumberingFormat format, string styleId)
        {
            _abstractNum = format.AbstractNum;
            _curLevel = _abstractNum.Elements<W.Level>().Where(l => l.ParagraphStyleIdInLevel?.Val == styleId).FirstOrDefault();
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets start number.
        /// </summary>
        public int Start
        {
            get => _curLevel.StartNumberingValue.Val;
        }

        /// <summary>
        /// Gets number style.
        /// </summary>
        public W.NumberFormatValues Style
        {
            get => _curLevel.NumberingFormat.Val;
        }

        /// <summary>
        /// Gets number format text.
        /// </summary>
        public string Format
        {
            get => _curLevel.LevelText.Val;
        }
        #endregion

        #region Internal Properties
        internal W.AbstractNum AbstractNum => _abstractNum;
        #endregion

    }
}
