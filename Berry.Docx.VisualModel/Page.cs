using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Berry.Docx.VisualModel.Documents;

namespace Berry.Docx.VisualModel
{
    public class Page
    {
        private readonly double _width = 0;
        private readonly double _availableWidth = 0;
        private readonly double _availableHeight = 0;
        private readonly double _height = 0;
        private readonly Margin _margin;
        private readonly double _charSpace = 0;
        private readonly double _lineSpace = 0;
        private readonly Berry.Docx.DocGridType _gridType;
        private readonly List<Paragraph> _paragraphs;

        private double _curHeight = 0;

        public Page(Berry.Docx.Document doc, Berry.Docx.Section section)
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
            _margin = new Margin(leftMar, topMar, rightMar, bottomMar);
            _gridType = pageSetup.DocGrid;
            _paragraphs = new List<Paragraph>();
        }

        public double Width => _width;

        public double Height => _height;

        public Margin Margin => _margin;

        public double CharSpace => _charSpace;

        public double LineSpace => _lineSpace;

        public List<Paragraph> Paragraphs => _paragraphs;

        public bool TryAppend(Berry.Docx.Documents.Paragraph p, ref int lineNumber)
        {
            Paragraph paragraph = new Paragraph(p, _availableWidth, _charSpace, _lineSpace, _gridType);
            var lines = paragraph.GenerateLines();
            if(_paragraphs.Count > 0)
            {
                var lastP = _paragraphs.Last();
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
                if(_curHeight + line.Height > _availableHeight)
                {
                    if(paragraph.Lines.Count > 0) _paragraphs.Add(paragraph);
                    return false;
                }
                if (line.HasPageBreak)
                {
                    paragraph.Lines.Add(line);
                    _paragraphs.Add(paragraph);
                    lineNumber++;
                    return false;
                }
                paragraph.Lines.Add(line);
                _curHeight += line.Height;
                lineNumber++;
            }
            if(paragraph.Lines.Count > 0) _paragraphs.Add(paragraph);
            return true;
        }
        
    }
}
