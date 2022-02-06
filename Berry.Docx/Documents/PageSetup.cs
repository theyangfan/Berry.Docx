using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    public class PageSetup
    {
        private W.SectionProperties _sectPr = null;
        private W.DocGrid _docGrid = null;
        internal PageSetup(Document doc, W.SectionProperties sectPr)
        {
            _sectPr = sectPr;
            _docGrid = sectPr.GetFirstChild<W.DocGrid>();
            if (_docGrid == null)
                _docGrid = new W.DocGrid();
            sectPr.AddChild(_docGrid);
        }


        /// <summary>
        /// 网格类型
        /// </summary>
        public DocGridType DocGridType
        {
            get
            {
                if (_docGrid.Type == null) return DocGridType.None;
                if (_docGrid.Type == W.DocGridValues.Lines)
                    return DocGridType.Lines;
                else if (_docGrid.Type == W.DocGridValues.LinesAndChars)
                    return DocGridType.LinesAndChars;
                else if (_docGrid.Type == W.DocGridValues.SnapToChars)
                    return DocGridType.SnapToChars;
                else
                    return DocGridType.None;
            }
            set
            {
                if (value == DocGridType.Lines)
                    _docGrid.Type = W.DocGridValues.Lines;
                else if (value == DocGridType.LinesAndChars)
                    _docGrid.Type = W.DocGridValues.LinesAndChars;
                else if (value == DocGridType.SnapToChars)
                    _docGrid.Type = W.DocGridValues.SnapToChars;
                else
                    _docGrid.Type = null;
            }
        }

        public int CharsCount
        {
            get
            {
                return 1/4096;
            }
        }


    }
}
