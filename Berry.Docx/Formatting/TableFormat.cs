using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;

namespace Berry.Docx.Formatting
{
    public class TableFormat
    {
        private readonly Document _doc;
        private readonly Table _table;
        private readonly W.Table _xtable;

        public TableFormat(Document doc, Table table)
        {
            _doc = doc;
            _table = table;
            _xtable = table.XElement;
        }

        /// <summary>
        /// Specifies that the first row format shall be applied to the table.
        /// </summary>
        public bool FirstRowEnabled
        {
            get
            {
                W.TableProperties tblPr = _xtable.GetFirstChild<W.TableProperties>();
                if(tblPr?.TableLook?.FirstRow != null)
                {
                    return tblPr.TableLook.FirstRow;
                }
                return true;
            }
            set
            {
                
                if(_xtable.GetFirstChild<W.TableProperties>() == null)
                {
                    _xtable.AddChild(new W.TableProperties());
                }
                W.TableProperties tblPr = _xtable.GetFirstChild<W.TableProperties>();
                if(tblPr.TableLook == null)
                {
                    tblPr.TableLook = new W.TableLook();
                }
                tblPr.TableLook.FirstRow = value;
            }
        }

        /// <summary>
        /// Specifies that the last row format shall be applied to the table.
        /// </summary>
        public bool LastRowEnabled
        {
            get
            {
                W.TableProperties tblPr = _xtable.GetFirstChild<W.TableProperties>();
                if (tblPr?.TableLook?.LastRow != null)
                {
                    return tblPr.TableLook.LastRow;
                }
                return false;
            }
            set
            {

                if (_xtable.GetFirstChild<W.TableProperties>() == null)
                {
                    _xtable.AddChild(new W.TableProperties());
                }
                W.TableProperties tblPr = _xtable.GetFirstChild<W.TableProperties>();
                if (tblPr.TableLook == null)
                {
                    tblPr.TableLook = new W.TableLook();
                }
                tblPr.TableLook.LastRow = value;
            }
        }

        /// <summary>
        /// Specifies that the first column format shall be applied to the table.
        /// </summary>
        public bool FirstColumnEnabled
        {
            get
            {
                W.TableProperties tblPr = _xtable.GetFirstChild<W.TableProperties>();
                if (tblPr?.TableLook?.FirstColumn != null)
                {
                    return tblPr.TableLook.FirstColumn;
                }
                return true;
            }
            set
            {

                if (_xtable.GetFirstChild<W.TableProperties>() == null)
                {
                    _xtable.AddChild(new W.TableProperties());
                }
                W.TableProperties tblPr = _xtable.GetFirstChild<W.TableProperties>();
                if (tblPr.TableLook == null)
                {
                    tblPr.TableLook = new W.TableLook();
                }
                tblPr.TableLook.FirstColumn = value;
            }
        }

        /// <summary>
        /// Specifies that the last column format shall be applied to the table.
        /// </summary>
        public bool LastColumnEnabled
        {
            get
            {
                W.TableProperties tblPr = _xtable.GetFirstChild<W.TableProperties>();
                if (tblPr?.TableLook?.LastColumn != null)
                {
                    return tblPr.TableLook.LastColumn;
                }
                return false;
            }
            set
            {

                if (_xtable.GetFirstChild<W.TableProperties>() == null)
                {
                    _xtable.AddChild(new W.TableProperties());
                }
                W.TableProperties tblPr = _xtable.GetFirstChild<W.TableProperties>();
                if (tblPr.TableLook == null)
                {
                    tblPr.TableLook = new W.TableLook();
                }
                tblPr.TableLook.LastColumn = value;
            }
        }

        public TableBorders Borders => new TableBorders(_doc, _table);
    }
}
