using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Media;

namespace Berry.Docx.Visual.Field
{
    public class Character
    {
        #region Private Members
        private readonly char _value;
        private readonly FormattedText _text;
        private readonly double _width = 0;
        private readonly double _fontWidth = 0;
        private readonly double _height = 0;
        private readonly VerticalAlignment _vAlign = VerticalAlignment.Center;
        private HorizontalAlignment _hAlign = HorizontalAlignment.Left;
        #endregion

        #region Constructor
        internal Character(Berry.Docx.Field.Character character, double charSpace, double normalFontSize, Berry.Docx.DocGridType gridType)
        {
            _value = character.Val;
            System.Globalization.CultureInfo culture = System.Globalization.CultureInfo.CurrentCulture;
            FlowDirection dir = FlowDirection.LeftToRight;
            FontFamily font = new FontFamily(character.FontName);
            FontWeight fontWeight = FontWeights.Normal;
            if (character.Bold) fontWeight = FontWeights.Bold;
            FontStyle fontStyle = FontStyles.Normal;
            if (character.Italic) fontStyle = FontStyles.Italic;
            Typeface typeface = new Typeface(font, fontStyle, fontWeight, FontStretches.Normal);

            double size = character.FontSize.ToPixel();
            System.Drawing.Color color = character.TextColor.Val;
            Brush brush = new SolidColorBrush(Color.FromRgb(color.R, color.G, color.B));
#if NET462_OR_GREATER
            double dpi = 1.0;
            if (Application.Current?.MainWindow != null)
                dpi = VisualTreeHelper.GetDpi(Application.Current.MainWindow).PixelsPerDip;
            else
                dpi = System.Drawing.Graphics.FromHwnd(IntPtr.Zero).DpiX / 96;
            _text = new FormattedText(character.Val.ToString(), culture, dir, typeface, size, brush, dpi);
#else
            _text = new FormattedText(character.Val.ToString(), culture, dir, typeface, size, brush);
#endif

            // 空格
            if (character.Val == 0x20)
            {
#if NET462_OR_GREATER
                _fontWidth = new FormattedText(".", culture, dir, typeface, size, brush, dpi).Width;
#else
                _fontWidth = new FormattedText(".", culture, dir, typeface, size, brush).Width;
#endif
            }
            else
            {
                _fontWidth = _text.Width;
            }

            if (character.SnapToGrid)
            {
                if (gridType == DocGridType.LinesAndChars)
                {
                    _width = charSpace + (character.FontSize - normalFontSize).ToPixel();
                }
                else if (gridType == DocGridType.SnapToChars)
                {
                    if (_fontWidth < charSpace)
                    {
                        _width = charSpace;
                    }
                    else
                    {
                        _width = Math.Ceiling(_fontWidth / charSpace) * charSpace;
                    }
                }
                else
                {
                    _width = _fontWidth;
                }
            }
            else
            {
                _width = _fontWidth;
            }

            _height = _text.Height;

            _vAlign = VerticalAlignment.Bottom;
            if (character.SnapToGrid && gridType == DocGridType.SnapToChars)
                _hAlign = HorizontalAlignment.Center;
            else
                _hAlign = HorizontalAlignment.Left;
        }
        #endregion

        #region Public Properties
        public char Val => _value;
        public FormattedText FormattedText => _text;

        public double Width => _width;

        public double Height => _height;

        public HorizontalAlignment HorizontalAlignment
        {
            get => _hAlign;
            internal set => _hAlign = value;
        }

        public VerticalAlignment VerticalAlignment => _vAlign;
        #endregion
    }
}
