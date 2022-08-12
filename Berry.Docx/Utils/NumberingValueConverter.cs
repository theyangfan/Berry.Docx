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

        private static Tuple<int, string>[] upperRomanSymbols =
        {
            new Tuple<int, string>(1000, "M"),
            new Tuple<int, string>(900, "CM"),
            new Tuple<int, string>(500, "D"),
            new Tuple<int, string>(400, "CD"),
            new Tuple<int, string>(100, "C"),
            new Tuple<int, string>(90, "XC"),
            new Tuple<int, string>(50, "L"),
            new Tuple<int, string>(40, "XL"),
            new Tuple<int, string>(10, "X"),
            new Tuple<int, string>(9, "IX"),
            new Tuple<int, string>(5, "V"),
            new Tuple<int, string>(4, "IV"),
            new Tuple<int, string>(1, "I")
        };

        private static Tuple<int, string>[] lowerRomanSymbols =
        {
            new Tuple<int, string>(1000, "m"),
            new Tuple<int, string>(900, "cm"),
            new Tuple<int, string>(500, "d"),
            new Tuple<int, string>(400, "cd"),
            new Tuple<int, string>(100, "c"),
            new Tuple<int, string>(90, "xc"),
            new Tuple<int, string>(50, "l"),
            new Tuple<int, string>(40, "xl"),
            new Tuple<int, string>(10, "x"),
            new Tuple<int, string>(9, "ix"),
            new Tuple<int, string>(5, "v"),
            new Tuple<int, string>(4, "iv"),
            new Tuple<int, string>(1, "i")
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
                int value = pair.Item1;
                string symbol = pair.Item2;
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
                int value = pair.Item1;
                string symbol = pair.Item2;
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
