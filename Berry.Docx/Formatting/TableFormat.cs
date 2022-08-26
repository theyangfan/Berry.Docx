using System;
using System.Collections.Generic;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represents the table format.
    /// </summary>
    public class TableFormat
    {
        #region Private Members
        private readonly Document _doc;
        private readonly Table _table;
        private readonly W.Table _xtable;
        #endregion

        #region Constructors
        internal TableFormat(Document doc, Table table)
        {
            _doc = doc;
            _table = table;
            _xtable = table.XElement;
        }
        #endregion

        #region Public Properties
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

        /// <summary>
        /// Gets or sets the table horizontal alignment.
        /// </summary>
        public TableRowAlignment HorizontalAlignment
        {
            get
            {
                W.TableProperties tblPr = _xtable.GetFirstChild<W.TableProperties>();
                if(tblPr?.TableJustification != null)
                {
                    return tblPr.TableJustification.Val.Value.Convert<TableRowAlignment>();
                }
                return _table.GetStyle().WholeTable.HorizontalAlignment;
            }
            set
            {
                if (_xtable.GetFirstChild<W.TableProperties>() == null)
                {
                    _xtable.AddChild(new W.TableProperties());
                }
                W.TableProperties tblPr = _xtable.GetFirstChild<W.TableProperties>();
                tblPr.TableJustification = new W.TableJustification() { Val = value.Convert<W.TableRowAlignmentValues>() };
                foreach(TableRow row in _table.Rows)
                {
                    row.HorizontalAlignment = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the table is floating.
        /// </summary>
        public bool WrapTextAround
        {
            get
            {
                W.TableProperties tblPr = _xtable.GetFirstChild<W.TableProperties>();
                return tblPr?.TablePositionProperties != null;
            }
            set
            {
                if (_xtable.GetFirstChild<W.TableProperties>() == null)
                {
                    _xtable.AddChild(new W.TableProperties());
                }
                W.TableProperties tblPr = _xtable.GetFirstChild<W.TableProperties>();
                if (value)
                {
                    if(tblPr.TablePositionProperties == null)
                    {
                        // set initial properties
                        tblPr.TablePositionProperties = new W.TablePositionProperties()
                        {
                            LeftFromText = 180,
                            RightFromText = 180,
                            VerticalAnchor = W.VerticalAnchorValues.Text,
                            TablePositionY = 1
                        };
                    }
                }
                else
                {
                    tblPr.TablePositionProperties = null;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether repeat the first row as header row at the top of each page.
        /// </summary>
        public bool RepeatHeaderRow
        {
            get
            {
                W.TableRow row = _table.Rows[0].XElement as W.TableRow;
                W.TableHeader header = row.TableRowProperties.GetFirstChild<W.TableHeader>();
                if (header == null) return false;
                if (header.Val == null) return true;
                return header.Val.Value == W.OnOffOnlyValues.On;
            }
            set
            {
                W.TableRow row = _table.Rows[0].XElement as W.TableRow;
                if (value)
                {
                    if (row.TableRowProperties == null)
                    {
                        row.TableRowProperties = new W.TableRowProperties();
                    }
                    row.TableRowProperties.AddChild(new W.TableHeader());
                }
                else
                {
                    row.TableRowProperties?.GetFirstChild<W.TableHeader>()?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets the table borders.
        /// </summary>
        public TableBorders Borders => new TableBorders(_doc, _table);
        #endregion
    }
}
