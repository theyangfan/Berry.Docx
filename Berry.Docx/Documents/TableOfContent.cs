using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;
using Berry.Docx.Field;

namespace Berry.Docx.Documents
{
    public class TableOfContent
    {
        private Document _doc;
        private SdtBlock _sdt;
        private string _fieldCode = string.Empty;
        internal TableOfContent(Document doc, SdtBlock sdt, string fieldCode)
        {
            _doc = doc;
            _sdt = sdt;
            _fieldCode = fieldCode;
        }

        public int StartOutlineLevel
        {
            get
            {
                string outlineLevel = Regex.Match(_fieldCode, @"\d+\s*\-\s*\d+").Value;
                if(string.IsNullOrEmpty(outlineLevel)) return 0;
                int.TryParse(outlineLevel.Split('-').FirstOrDefault(), out int start);
                return start;
            }
        }

        public int EndOutlineLevel
        {
            get
            {
                string outlineLevel = Regex.Match(_fieldCode, @"\d+\s*\-\s*\d+").Value;
                if (string.IsNullOrEmpty(outlineLevel)) return 0;
                int.TryParse(outlineLevel.Split('-').LastOrDefault(), out int end);
                return end;
            }
        }

        public void Update()
        {

        }

        private void Reset(string code)
        {
            Paragraph tocP = new Paragraph(_doc);
            tocP.AppendText("目录");
            tocP.ApplyStyle(BuiltInStyle.TOC1);

            Paragraph tocBegin = new Paragraph(_doc);
            Paragraph tocEnd = new Paragraph(_doc);
            var fieldBegin = new FieldChar(_doc, FieldCharType.Begin);
            var fieldCode = new FieldCode(_doc, code);
            var fieldSeparate = new FieldChar(_doc, FieldCharType.Separate);
            var fieldEnd = new FieldChar(_doc, FieldCharType.End);
            tocBegin.ChildItems.Add(fieldBegin);
            tocBegin.ChildItems.Add(fieldCode);
            tocBegin.ChildItems.Add(fieldSeparate);
            tocEnd.ChildItems.Add(fieldEnd);

            _sdt.Content.ChildObjects.Add(tocBegin);
            _sdt.Content.ChildObjects.Add(tocEnd);
        }
    }
}
