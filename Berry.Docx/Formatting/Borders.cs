using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    public class Borders
    {
        private readonly Document _doc;
        private readonly W.Paragraph _ownerParagraph;
        private readonly W.Style _ownerStyle;

        internal Borders(Document doc, W.Paragraph paragraph)
        {
            _doc = doc;
            _ownerParagraph = paragraph;
        }

        internal Borders(Document doc, W.Style style)
        {
            _doc = doc;
            _ownerStyle = style;
        }

        public Border Top
        {
            get
            {
                if(_ownerParagraph != null)
                {
                    return new Border(_doc, _ownerParagraph, BorderType.Top);
                }
                else
                {
                    return new Border(_doc, _ownerStyle, BorderType.Top);
                }
            }
        }

        public Border Bottom
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return new Border(_doc, _ownerParagraph, BorderType.Bottom);
                }
                else
                {
                    return new Border(_doc, _ownerStyle, BorderType.Bottom);
                }
            }
        }

        public Border Left
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return new Border(_doc, _ownerParagraph, BorderType.Left);
                }
                else
                {
                    return new Border(_doc, _ownerStyle, BorderType.Left);
                }
            }
        }

        public Border Right
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return new Border(_doc, _ownerParagraph, BorderType.Right);
                }
                else
                {
                    return new Border(_doc, _ownerStyle, BorderType.Right);
                }
            }
        }

        public Border Between
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return new Border(_doc, _ownerParagraph, BorderType.Between);
                }
                else
                {
                    return new Border(_doc, _ownerStyle, BorderType.Between);
                }
            }
        }

        public Border Bar
        {
            get
            {
                if (_ownerParagraph != null)
                {
                    return new Border(_doc, _ownerParagraph, BorderType.Bar);
                }
                else
                {
                    return new Border(_doc, _ownerStyle, BorderType.Bar);
                }
            }
        }

        public void Clear()
        {
            if(Top.Style != BorderStyle.Nil || Top.Style != BorderStyle.None)
            {
                Top.Style = BorderStyle.None;
                Top.Color = ColorValue.Auto;
                Top.Width = 0;
            }
            if (Bottom.Style != BorderStyle.Nil || Bottom.Style != BorderStyle.None)
            {
                Bottom.Style = BorderStyle.None;
                Bottom.Color = ColorValue.Auto;
                Bottom.Width = 0;
            }
            if (Left.Style != BorderStyle.Nil || Left.Style != BorderStyle.None)
            {
                Left.Style = BorderStyle.None;
                Left.Color = ColorValue.Auto;
                Left.Width = 0;
            }
            if (Right.Style != BorderStyle.Nil || Right.Style != BorderStyle.None)
            {
                Right.Style = BorderStyle.None;
                Right.Color = ColorValue.Auto;
                Right.Width = 0;
            }
            if (Between.Style != BorderStyle.Nil || Between.Style != BorderStyle.None)
            {
                Between.Style = BorderStyle.None;
                Between.Color = ColorValue.Auto;
                Between.Width = 0;
            }
            if (Bar.Style != BorderStyle.Nil || Bar.Style != BorderStyle.None)
            {
                Bar.Style = BorderStyle.None;
                Bar.Color = ColorValue.Auto;
                Bar.Width = 0;
            }
        }
    }
}
