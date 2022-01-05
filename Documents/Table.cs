using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;
namespace Berry.Docx.Documents
{
    public class Table : DocumentObject
    {
        private Document _doc = null;
        private W.Table _table = null;
        public Table(Document doc, W.Table table) : base(doc, table)
        {
            _doc = doc;
            _table = table;
        }
        /// <summary>
        /// 表格单元格中所有的段落
        /// </summary>
        public List<Paragraph> Paragraphs
        {
            get
            {
                List<Paragraph> paras = new List<Paragraph>();
                foreach (W.Paragraph p in _table.Descendants<W.Paragraph>())
                    paras.Add(new Paragraph(_doc, p));
                return paras;
            }
        }
    }
}
