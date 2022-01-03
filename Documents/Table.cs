using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;
namespace Berry.Docx.Documents
{
    public class Table : DocumentObject
    {
        private W.Table _table = null;
        public Table(W.Table table) : base(table)
        {
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
                    paras.Add(new Paragraph(p));
                return paras;
            }
        }
    }
}
