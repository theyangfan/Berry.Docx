using Berry.Docx.Collections;
using System;
using System.Collections.Generic;
using System.Text;
using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    public class TableCell : DocumentElement
    {
        private Document _doc;
        private W.TableCell _cell;
        public TableCell(Document doc)
            :this(doc, new W.TableCell())
        {

        }

        internal TableCell(Document doc, W.TableCell cell):base(doc, cell)
        {
            _doc = doc;
            _cell = cell;
        }

        public override DocumentObjectType DocumentObjectType => DocumentObjectType.TableCell;

        public override DocumentObjectCollection ChildObjects => Paragraphs;

        public ParagraphCollection Paragraphs => new ParagraphCollection(_cell, GetParagraphs());

        private IEnumerable<Paragraph> GetParagraphs()
        {
            foreach(W.Paragraph p in _cell.Elements<W.Paragraph>())
            {
                yield return new Paragraph(_doc, p);
            }
        }
    }
}
