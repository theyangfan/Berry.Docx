using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    internal class NumberValueConverter
    {
        private static List<string> cnSymbols = new List<string>() { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
        private static Dictionary<int, string> unitSymbols = new Dictionary<int, string>()
        {
            { 2, "十" },
            { 3, "百" },
            { 4, "千" },
            { 5, "万" },
            { 6, "十" },
            { 7, "百" },
            { 8, "千" },
            { 9, "亿" },
            { 10, "十" },
            { 11, "百" },
            { 12, "千" },
            { 13, "万" },
        };
#if NEt35
#else
#endif
        private static Dictionary<int, string> upperRomanSymbols = new Dictionary<int, string>()
        {
            { 1000, "M" },
            { 900, "CM" },
            { 500, "D" },
            { 400, "CD" },
            { 100, "C" },
            { 90, "XC" },
            { 50, "L" },
            { 40, "XL" },
            { 10, "X" },
            { 9, "IX" },
            { 5, "V" },
            { 4, "IV" },
            { 1, "I" }
        };

        private static Dictionary<int, string> lowerRomanSymbols = new Dictionary<int, string>()
        {
            { 1000, "m" },
            { 900, "cm" },
            { 500, "d" },
            { 400, "cd" },
            { 100, "c" },
            { 90, "xc" },
            { 50, "l" },
            { 40, "xl" },
            { 10, "x" },
            { 9, "ix" },
            { 5, "v" },
            { 4, "iv" },
            { 1, "i" }
        };

        public static string IntToChineseCounting(int num)
        {
            if (num < 0) return string.Empty;
            StringBuilder symbols = new StringBuilder();
            string strNum = num.ToString();
            while(strNum.Length > 0)
            {
                if(strNum.Length == 1)
                {
                    if (strNum[0] != '0' || symbols.Length == 0) symbols.Append(cnSymbols[strNum.ToInt()]);
                    break;
                }
                else
                {
                    symbols.Append(cnSymbols[strNum[0] - 48]);
                    if (strNum[0] != '0' || strNum.Length == 5 || strNum.Length == 9)
                        symbols.Append(unitSymbols[strNum.Length]);
                    strNum = strNum.Remove(0, 1);
                }
            }
            string symbol = symbols.ToString().RxReplace("零+", "零").RxReplace("零亿", "亿").RxReplace("零万","万");
            if(symbol.Length > 1 && symbol.EndsWith("零")) symbol = symbol.Substring(0, symbol.Length - 1);
            return symbol;
        }

        public static string IntToUpperRoman(int num)
        {
            StringBuilder roman = new StringBuilder();
            foreach(var pair in upperRomanSymbols)
            {
                int value = pair.Key;
                string symbol = pair.Value;
                while(num >= value)
                {
                    num -= value;
                    roman.Append(symbol);
                }
                if(num == 0)
                {
                    break;
                }
            }
            return roman.ToString();
        }

        public static string IntToLowerRoman(int num)
        {
            StringBuilder roman = new StringBuilder();
            foreach (var pair in lowerRomanSymbols)
            {
                int value = pair.Key;
                string symbol = pair.Value;
                while (num >= value)
                {
                    num -= value;
                    roman.Append(symbol);
                }
                if (num == 0)
                {
                    break;
                }
            }
            return roman.ToString();
        }
    }
}
