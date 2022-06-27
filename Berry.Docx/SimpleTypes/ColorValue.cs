using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

namespace Berry.Docx
{
    public class ColorValue
    {
        private Color _color = Color.Empty;
        private bool _auto = false;

        public static ColorValue Auto = new ColorValue() { IsAuto = true };
       
        public ColorValue() { }

        public ColorValue(string rgb)
        {
            _color = ColorConverter.FromHex(rgb);
            _auto = rgb == "auto";
        }

        public ColorValue(Color color)
        {
            _color= color;
        }
        public ColorValue(ColorValue source)
        {
            _color = source.Val;
        }

        public bool IsAuto
        {
            get => _auto;
            set => _auto = value;
        }

        public Color Val
        {
            get => _color;
            set => _color = value;
        }

        public static implicit operator Color(ColorValue value)
        {
            return value.Val;
        }

        public static implicit operator ColorValue(Color value)
        {
            return new ColorValue(value);
        }

        public static implicit operator ColorValue(string rgb)
        {
            return new ColorValue(rgb);
        }

        public override string ToString()
        {
            if (_auto) return "auto";
            else return ColorConverter.ToHex(_color);
        }
    }
}
