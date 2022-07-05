using System;
using System.Collections.Generic;
using System.Text;

namespace Berry.Docx
{
    internal class BuiltInStyleNameFormatter
    {
        private static Dictionary<string, string> STYLE_NAMES = new Dictionary<string, string>()
        {
            { "正文", "normal" },
            { "标题 1", "heading 1" },
            { "标题 2", "heading 2" },
            { "标题 3", "heading 3" },
            { "标题 4", "heading 4" },
            { "标题 5", "heading 5" },
            { "标题 6", "heading 6" },
            { "标题 7", "heading 7" },
            { "标题 8", "heading 8" },
            { "标题 9", "heading 9" },
            { "目录 1", "toc 1" },
            { "目录 2", "toc 2" },
            { "目录 3", "toc 3" },
            { "目录 4", "toc 4" },
            { "目录 5", "toc 5" },
            { "目录 6", "toc 6" },
            { "目录 7", "toc 7" },
            { "目录 8", "toc 8" },
            { "目录 9", "toc 9" }
        };

        public static string NameToBuiltInString(string styleName)
        {
            styleName = styleName.ToLower();
            if (STYLE_NAMES.ContainsKey(styleName))
            {
                return STYLE_NAMES[styleName];
            }
            return styleName;
        }
    }
}
