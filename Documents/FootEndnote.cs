using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OW = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    public class FootEndnote
    {
        private Document _doc = null;
        private OW.Footnote _footnote;
        private OW.Endnote _endnote;
        public FootEndnote(Document doc, OW.Footnote footnote)
        {
            _doc = doc;
            _footnote = footnote;
        }
        public FootEndnote(Document doc, OW.Endnote endnote)
        {
            _doc = doc;
            _endnote = endnote;
        }
        /// <summary>
        /// 获取脚注段落
        /// </summary>
        /// <returns></returns>
        public List<Paragraph> GetFootnoteParagrahs()
        {
            List<Paragraph> para = new List<Paragraph>();
            if (_footnote.Elements<OW.Paragraph>().Any())
            {
                foreach (OW.Paragraph p in _footnote.Elements<OW.Paragraph>())
                {
                    Paragraph myPara = new Paragraph(_doc, p);
                    para.Add(myPara);
                }
            }
            return para;
        }
        /// <summary>
        /// 获取尾注段落
        /// </summary>
        /// <returns></returns>
        public List<Paragraph> GetEndnoteParagrahs()
        {
            List<Paragraph> para = new List<Paragraph>();
            if (_endnote.Elements<OW.Paragraph>().Any())
            {
                foreach (OW.Paragraph p in _endnote.Elements<OW.Paragraph>())
                {
                    Paragraph myPara = new Paragraph(_doc, p);
                    para.Add(myPara);
                }
            }
            return para;
        }

        /// <summary>
        /// 清除脚注分隔符
        /// </summary>
        public void ClearFSeparator()
        {
            if (_footnote.Type != null && _footnote.Type == OW.FootnoteEndnoteValues.Separator)
            {
                _footnote.RemoveAllChildren<OW.Paragraph>();
            }
        }
        /// <summary>
        /// 测试
        /// </summary>
        public void Test()
        {
            Console.WriteLine("test");
        }
        /// <summary>
        /// 清除脚注延续分隔符
        /// </summary>
        public void ClearContinuationFSeparator()
        {
            if (_footnote.Type != null && _footnote.Type == OW.FootnoteEndnoteValues.ContinuationSeparator)
            {
                _footnote.RemoveAllChildren<OW.Paragraph>();
            }
        }
        /// <summary>
        /// 清除尾注分隔符
        /// </summary>
        public void ClearESeparator()
        {
            if (_endnote.Type != null && _endnote.Type == OW.FootnoteEndnoteValues.Separator)
            {
                _endnote.RemoveAllChildren<OW.Paragraph>();
            }
        }
        /// <summary>
        /// 清除尾注延续分隔符
        /// </summary>
        public void ClearContinuationESeparator()
        {

            if (_endnote.Type != null && _endnote.Type == OW.FootnoteEndnoteValues.ContinuationSeparator)
            {
                _endnote.RemoveAllChildren<OW.Paragraph>();
            }
        }
        /// <summary>
        /// 获取脚注分隔符/延续分隔符段落
        /// </summary>
        public List<Paragraph> FSeparatorParagraphs()
        {
            List<Paragraph> p = new List<Paragraph>();
            if (_footnote.Elements<OW.Paragraph>().Any())
            {
                OW.Paragraph paragraph = _footnote.Elements<OW.Paragraph>().First();
                if (paragraph.Descendants<OW.SeparatorMark>().Any())
                {
                    Paragraph myPara = new Paragraph(_doc, paragraph);
                    
                    p.Add(myPara);
                }
                if (paragraph.Descendants<OW.ContinuationSeparatorMark>().Any())
                {
                    Paragraph myPara = new Paragraph(_doc, paragraph);
                    
                    p.Add(myPara);
                }
            }
            return p;
        }
        /// <summary>
        /// 获取尾注分隔符/延续分隔符段落
        /// </summary>
        public List<Paragraph> ESeparatorParagraphs()
        {
            List<Paragraph> p = new List<Paragraph>();
            if (_endnote.Elements<OW.Paragraph>().Any())
            {
                OW.Paragraph paragraph = _endnote.Elements<OW.Paragraph>().First();
                if (paragraph.Descendants<OW.SeparatorMark>().Any())
                {
                    Paragraph myPara = new Paragraph(_doc, paragraph);
                    p.Add(myPara);
                }
                if (paragraph.Descendants<OW.ContinuationSeparatorMark>().Any())
                {
                    Paragraph myPara = new Paragraph(_doc, paragraph);
                    p.Add(myPara);
                }
            }
            return p;
        }
    }
}
