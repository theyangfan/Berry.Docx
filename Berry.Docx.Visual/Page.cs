using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using BD = Berry.Docx.Documents;
using Berry.Docx.Visual.Documents;
using Berry.Docx.Visual.Field;

namespace Berry.Docx.Visual
{
    public class Page
    {
        #region Private Members
        private readonly double _width = 0;
        private readonly double _availableWidth = 0;
        private readonly double _availableHeight = 0;
        private readonly double _height = 0;
        private readonly Margin _padding = new Margin(0,0,0,0);
        private readonly double _charSpace = 0;
        private readonly double _lineSpace = 0;
        private readonly Berry.Docx.DocGridType _gridType;
        private readonly List<DocumentItem> _items;

        private double _curHeight = 0;
        #endregion

        #region Constructor
        internal Page(Berry.Docx.Document doc, Berry.Docx.Section section)
        {
            var pageSetup = section.PageSetup;
            _width = pageSetup.PageWidth.ToPixel();
            _height = pageSetup.PageHeight.ToPixel();
            double topMar = pageSetup.TopMargin.ToPixel();
            double bottomMar = pageSetup.BottomMargin.ToPixel();
            double leftMar = pageSetup.LeftMargin.ToPixel();
            double rightMar = pageSetup.RightMargin.ToPixel();
            _charSpace = pageSetup.CharPitch.ToPixel();
            _lineSpace = pageSetup.LinePitch.ToPixel();

            _availableWidth = _width - leftMar - rightMar;
            _availableHeight = _height - topMar - bottomMar;
            _padding = new Margin(leftMar, topMar, rightMar, bottomMar);
            _gridType = pageSetup.DocGrid;
            _items = new List<DocumentItem>();
        }
        #endregion

        #region Public Properties
        public double Width => _width;

        public double Height => _height;

        public Margin Padding => _padding;

        public double CharSpace => _charSpace;

        public double LineSpace => _lineSpace;

        public List<DocumentItem> ChildItems => _items;

        #endregion

        #region Internal Methods
        internal bool TryAppend(BD.Paragraph p, ref int lineNumber)
        {
            Paragraph paragraph = new Paragraph(p, _availableWidth, _charSpace, _lineSpace, _gridType);
            var lines = paragraph.GenerateLines();
            if(_items.Count > 0 && _items.Last() is Paragraph)
            {
                var lastP = (Paragraph)_items.Last();
                double margin = Math.Max(lastP.Margin.Bottom, paragraph.Margin.Top);
                lastP.Margin.Bottom = 0;
                paragraph.Margin.Top = margin;
            }
            if(_curHeight + paragraph.Margin.Top > _availableHeight)
            {
                return false;
            }
            _curHeight += paragraph.Margin.Top;
            int count = lines.Count;
            for (int i = lineNumber; i < count; i++)
            {
                var line = lines[i];
                if(line.ChildItems[0].GetType() != typeof(Picture) || _items.Count > 0)
                {
                    if (_curHeight + line.Height + line.Margin.Top + line.Margin.Bottom > _availableHeight)
                    {
                        if (paragraph.Lines.Count > 0) _items.Add(paragraph);
                        return false;
                    }
                }
                if (line.EndsWithPageBreak)
                {
                    paragraph.Lines.Add(line);
                    _items.Add(paragraph);
                    lineNumber++;
                    return false;
                }
                paragraph.Lines.Add(line);
                _curHeight += line.Height + line.Margin.Top + line.Margin.Bottom;
                lineNumber++;
            }
            if(paragraph.Lines.Count > 0) _items.Add(paragraph);
            return true;
        }

        internal bool TryAppend(BD.Table tbl)
        {
            Table table = new Table(tbl, _charSpace, _lineSpace, _gridType);
            if (_items.Count > 0 && _curHeight + table.Height > _availableHeight) return false;
            _items.Add(table);
            _curHeight += table.Height;
            return true;
        }
        #endregion
    }
}
