using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    public class TableRegionStyle
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Style _style;
        private readonly TableRegionType _region;
        private readonly CharacterFormat _cFormat;
        private readonly ParagraphFormat _pFormat;
        #endregion

        #region Constructors
        internal TableRegionStyle(Document doc, W.Style style, TableRegionType region)
        {
            _doc = doc;
            _style = style;
            _region = region;
            _cFormat = new CharacterFormat(doc, style, region);
            _pFormat = new ParagraphFormat(doc, style, region);
        }
        #endregion

        #region Public Properties
        public CharacterFormat CharacterFormat => _cFormat;

        public ParagraphFormat ParagraphFormat => _pFormat;
        #endregion

    }

    internal enum TableRegionType
    {
        WholeTable = 0,
        FirstRow = 1,
        LastRow = 2,
        FirstColumn = 3,
        LastColumn = 4
    }
}
