using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Documents;

namespace Berry.Docx.Formatting
{
    internal class TablePropertiesHolder
    {
        private readonly W.TableCell _cell;
        private readonly W.TableRow _row;
        private readonly W.Table _table;
        private readonly W.Style _style;
        private readonly TableRegionType _region;

        public TablePropertiesHolder(TableCell cell)
        {
            _cell = cell.XElement;
        }

        public TablePropertiesHolder(TableRow row)
        {
            _row = row.XElement;
        }

        public TablePropertiesHolder(Table table)
        {
            _table = table.XElement;
        }

        public TablePropertiesHolder(Style style, TableRegionType region)
        {
            _style = style.XElement;
            _region = region;
        }

        #region Table, Cell & Style
        public ColorValue Background
        {
            get
            {
                W.Shading shd = null;
                if (_cell != null)
                {
                    shd = _cell.TableCellProperties?.Shading;
                }
                else if (_table != null)
                {
                    shd = _table.GetFirstChild<W.TableProperties>()?.GetFirstChild<W.Shading>();
                }
                else if (_style != null)
                {
                    if (_region == TableRegionType.WholeTable)
                    {
                        shd = _style.StyleTableCellProperties?.Shading;
                    }
                    else
                    {
                        var type = _region.Convert<W.TableStyleOverrideValues>();
                        var tblStylePr = _style.Elements<W.TableStyleProperties>()
                            .Where(t => t.Type == type).FirstOrDefault();
                        shd = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.Shading;
                    }
                }
                if (shd?.Fill == null) return null;
                return shd.Fill.Value;
            }
            set
            {
                var shd = new W.Shading() { Val = W.ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                if (_cell != null)
                {
                    if (_cell.TableCellProperties == null)
                        _cell.TableCellProperties = new W.TableCellProperties();
                    if (_cell.TableCellProperties.Shading == null)
                        _cell.TableCellProperties.Shading = shd;
                    _cell.TableCellProperties.Shading.Fill = value.ToString();
                }
                else if (_table != null)
                {
                    if (_table.GetFirstChild<W.TableProperties>() == null)
                        _table.AddChild(new W.TableProperties());
                    if (_table.GetFirstChild<W.TableProperties>().GetFirstChild<W.Shading>() == null)
                        _table.GetFirstChild<W.TableProperties>().AddChild(shd);
                    _table.GetFirstChild<W.TableProperties>().GetFirstChild<W.Shading>().Fill = value.ToString();
                }
                else if (_style != null)
                {
                    if (_region == TableRegionType.WholeTable)
                    {
                        if (_style.StyleTableCellProperties == null)
                            _style.StyleTableCellProperties = new W.StyleTableCellProperties();
                        if (_style.StyleTableCellProperties.Shading == null)
                            _style.StyleTableCellProperties.Shading = shd;
                        _style.StyleTableCellProperties.Shading.Fill = value.ToString();
                    }
                    else
                    {
                        var type = _region.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        var tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.TableStyleConditionalFormattingTableCellProperties == null)
                            tblStylePr.TableStyleConditionalFormattingTableCellProperties = new W.TableStyleConditionalFormattingTableCellProperties();
                        if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.Shading == null)
                            tblStylePr.TableStyleConditionalFormattingTableCellProperties.Shading = shd;
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.Shading.Fill = value.ToString();
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the horizontal alignment.
        /// </summary>
        public EnumValue<TableRowAlignment> HorizontalAlignment
        {
            get
            {
                W.TableJustification jc = null;
                if(_row != null)
                {
                    jc = _row.TableRowProperties?.GetFirstChild<W.TableJustification>();
                }
                else if(_table != null)
                {
                    jc = _table.GetFirstChild<W.TableProperties>()?.GetFirstChild<W.TableJustification>();
                }
                else if(_style != null)
                {
                    if (_region == TableRegionType.WholeTable)
                    {
                        jc = _style.StyleTableProperties?.TableJustification;
                    }
                    else
                    {
                        var type = _region.Convert<W.TableStyleOverrideValues>();
                        var tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        jc = tblStylePr?.TableStyleConditionalFormattingTableProperties?.TableJustification;
                    }
                }
                if (jc == null) return null;
                if (jc.Val == null) return TableRowAlignment.Left;
                return jc.Val.Value.Convert<TableRowAlignment>();
            }
            set
            {
                var jc = new W.TableJustification() { Val = value.Val.Convert<W.TableRowAlignmentValues>() };
                if(_row != null)
                {
                    if(_row.TableRowProperties == null)
                        _row.TableRowProperties = new W.TableRowProperties();
                    _row.TableRowProperties.AddChild(jc);
                }
                else if (_table != null)
                {
                    if (_table.GetFirstChild<W.TableProperties>() == null)
                        _table.AddChild(new W.TableProperties());
                    _table.GetFirstChild<W.TableProperties>().AddChild(jc);
                }
                else if(_style != null)
                {
                    if (_region == TableRegionType.WholeTable)
                    {
                        if (_style.StyleTableProperties == null)
                            _style.StyleTableProperties = new W.StyleTableProperties();
                        _style.StyleTableProperties.TableJustification = jc;
                        if (_style.TableStyleConditionalFormattingTableRowProperties == null)
                            _style.TableStyleConditionalFormattingTableRowProperties = new W.TableStyleConditionalFormattingTableRowProperties();
                        _style.TableStyleConditionalFormattingTableRowProperties.AddChild(jc);
                    }
                    else if (_region == TableRegionType.FirstRow || _region == TableRegionType.LastRow)
                    {
                        var type = _region.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.TableStyleConditionalFormattingTableProperties == null)
                            tblStylePr.TableStyleConditionalFormattingTableProperties = new W.TableStyleConditionalFormattingTableProperties();
                        tblStylePr.TableStyleConditionalFormattingTableProperties.TableJustification = jc;
                        if (tblStylePr.TableStyleConditionalFormattingTableRowProperties == null)
                            tblStylePr.TableStyleConditionalFormattingTableRowProperties = new W.TableStyleConditionalFormattingTableRowProperties();
                        tblStylePr.TableStyleConditionalFormattingTableRowProperties.AddChild(jc);
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether allow row to break across pages.
        /// </summary>
        public BooleanValue AllowBreakAcrossPages
        {
            get
            {
                W.CantSplit cantSplit = null;
                if (_row != null)
                {
                    cantSplit = _row.TableRowProperties?.GetFirstChild<W.CantSplit>();
                }
                else if (_style != null && _region == TableRegionType.WholeTable)
                {
                    cantSplit = _style.TableStyleConditionalFormattingTableRowProperties?.GetFirstChild<W.CantSplit>();
                }
                if (cantSplit == null) return null;
                if (cantSplit.Val == null) return false;
                return cantSplit.Val.Value == W.OnOffOnlyValues.Off;
            }
            set
            {
                if(_row != null)
                {
                    if (value)
                    {
                        _row.TableRowProperties?.RemoveAllChildren<W.CantSplit>();
                    }
                    else
                    {
                        if (_row.TableRowProperties == null)
                            _row.TableRowProperties = new W.TableRowProperties();
                        _row.TableRowProperties.AddChild(new W.CantSplit());
                    }
                }
                else if (_style != null && _region == TableRegionType.WholeTable)
                {
                    if (value)
                    {
                        _style.TableStyleConditionalFormattingTableRowProperties?.RemoveAllChildren<W.CantSplit>();
                    }
                    else
                    {
                        if (_style.TableStyleConditionalFormattingTableRowProperties == null)
                            _style.TableStyleConditionalFormattingTableRowProperties = new W.TableStyleConditionalFormattingTableRowProperties();
                        _style.TableStyleConditionalFormattingTableRowProperties.AddChild(new W.CantSplit());
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the table cell vertical alignment.
        /// </summary>
        public EnumValue<TableCellVerticalAlignment> VerticalCellAlignment
        {
            get
            {
                W.TableCellVerticalAlignment vAlign = null;
                if (_cell != null)
                {
                    vAlign = _cell.TableCellProperties?.TableCellVerticalAlignment;
                }
                else if (_style != null)
                {
                    if (_region == TableRegionType.WholeTable)
                    {
                        vAlign = _style.StyleTableCellProperties?.TableCellVerticalAlignment;
                    }
                    else
                    {
                        var type = _region.Convert<W.TableStyleOverrideValues>();
                        var tblStylePr = _style.Elements<W.TableStyleProperties>()
                            .Where(t => t.Type == type).FirstOrDefault();
                        vAlign = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellVerticalAlignment;
                    }
                }
                if (vAlign == null) return null;
                if (vAlign.Val == null) return TableCellVerticalAlignment.Top;
                return vAlign.Val.Value.Convert<TableCellVerticalAlignment>();
            }
            set
            {
                var vAlign = new W.TableCellVerticalAlignment() { Val = value.Val.Convert<W.TableVerticalAlignmentValues>() };
                if (_cell != null)
                {
                    if (_cell.TableCellProperties == null)
                        _cell.TableCellProperties = new W.TableCellProperties();
                    _cell.TableCellProperties.TableCellVerticalAlignment = vAlign;
                }
                else if (_style != null)
                {
                    if (_region == TableRegionType.WholeTable)
                    {
                        if (_style.StyleTableCellProperties == null)
                            _style.StyleTableCellProperties = new W.StyleTableCellProperties();
                        _style.StyleTableCellProperties.TableCellVerticalAlignment = vAlign;
                    }
                    else
                    {
                        var type = _region.Convert<W.TableStyleOverrideValues>();
                        if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                            _style.Append(new W.TableStyleProperties() { Type = type });
                        var tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                        if (tblStylePr.TableStyleConditionalFormattingTableCellProperties == null)
                            tblStylePr.TableStyleConditionalFormattingTableCellProperties = new W.TableStyleConditionalFormattingTableCellProperties();
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellVerticalAlignment = vAlign;
                    }
                }
            }
        }

        #endregion


        #region Only Cell
        public IntegerValue ColumnSpan
        {
            get
            {
                W.GridSpan gridSpan = null;
                if (_cell != null)
                {
                    gridSpan = _cell.TableCellProperties?.GridSpan;
                }
                if (gridSpan?.Val != null) return gridSpan.Val.Value;
                return null;
            }
            set
            {
                if(_cell != null)
                {
                    if (value <= 1 && _cell.TableCellProperties?.GridSpan != null)
                    {
                        _cell.TableCellProperties.GridSpan = null;
                        return;
                    }
                    if (_cell.TableCellProperties == null)
                        _cell.TableCellProperties = new W.TableCellProperties();
                    _cell.TableCellProperties.GridSpan = new W.GridSpan() { Val = value.Val };
                }
            }
        }

        public EnumValue<TableCellVerticalMergeType> VMerge
        {
            get
            {
                W.VerticalMerge vMerge = null;
                if (_cell != null)
                {
                    vMerge = _cell.TableCellProperties?.VerticalMerge;
                }
                if (vMerge == null) return null;
                if (vMerge.Val == null) return TableCellVerticalMergeType.Continue;
                return vMerge.Val.Value.Convert<TableCellVerticalMergeType>();
            }
            set
            {
                if(_cell != null)
                {
                    if (value == TableCellVerticalMergeType.None && _cell.TableCellProperties?.VerticalMerge != null)
                    {
                        _cell.TableCellProperties.VerticalMerge = null;
                        return;
                    }
                    if (_cell.TableCellProperties == null)
                        _cell.TableCellProperties = new W.TableCellProperties();
                    if (value == TableCellVerticalMergeType.Restart)
                        _cell.TableCellProperties.VerticalMerge = new W.VerticalMerge() { Val = W.MergedCellValues.Restart };
                    else if (value == TableCellVerticalMergeType.Continue)
                        _cell.TableCellProperties.VerticalMerge = new W.VerticalMerge();
                }
            }
        }
        #endregion

        #region Only Row
        public BooleanValue RepeatHeaderRow
        {
            get
            {
                W.TableHeader header = null;
                if(_row != null)
                {
                    header = _row.TableRowProperties?.GetFirstChild<W.TableHeader>();
                }
                if (header == null) return null;
                if (header.Val == null) return true;
                return header.Val.Value == W.OnOffOnlyValues.On;
            }
            set
            {
                if (_row != null)
                {
                    if (value)
                    {
                        if (_row.TableRowProperties == null)
                            _row.TableRowProperties = new W.TableRowProperties();
                        _row.TableRowProperties.AddChild(new W.TableHeader());
                    }
                    else
                    {
                        _row.TableRowProperties?.RemoveAllChildren<W.TableHeader>();
                    }
                }
            }
        }
        #endregion

        #region Only Table
        /// <summary>
        /// Specifies that the first row format shall be applied to the table.
        /// </summary>
        public BooleanValue FirstRowEnabled
        {
            get
            {
                W.TableLook look = null;
                if(_table != null)
                {
                    look = _table.GetFirstChild<W.TableProperties>()?.TableLook;
                }
                if (look?.FirstRow != null) return look.FirstRow.Value;
                return null;
            }
            set
            {
                if(_table != null)
                {
                    if (_table.GetFirstChild<W.TableProperties>() == null)
                        _table.AddChild(new W.TableProperties());
                    var tblPr = _table.GetFirstChild<W.TableProperties>();
                    if (tblPr.TableLook == null)
                        tblPr.TableLook = new W.TableLook();
                    tblPr.TableLook.FirstRow = value.Val;
                }
            }
        }

        /// <summary>
        /// Specifies that the last row format shall be applied to the table.
        /// </summary>
        public BooleanValue LastRowEnabled
        {
            get
            {
                W.TableLook look = null;
                if (_table != null)
                {
                    look = _table.GetFirstChild<W.TableProperties>()?.TableLook;
                }
                if (look?.LastRow != null) return look.LastRow.Value;
                return null;
            }
            set
            {
                if (_table != null)
                {
                    if (_table.GetFirstChild<W.TableProperties>() == null)
                        _table.AddChild(new W.TableProperties());
                    var tblPr = _table.GetFirstChild<W.TableProperties>();
                    if (tblPr.TableLook == null)
                        tblPr.TableLook = new W.TableLook();
                    tblPr.TableLook.LastRow = value.Val;
                }
            }
        }

        /// <summary>
        /// Specifies that the first column format shall be applied to the table.
        /// </summary>
        public BooleanValue FirstColumnEnabled
        {
            get
            {
                W.TableLook look = null;
                if (_table != null)
                {
                    look = _table.GetFirstChild<W.TableProperties>()?.TableLook;
                }
                if (look?.FirstColumn != null) return look.FirstColumn.Value;
                return null;
            }
            set
            {
                if (_table != null)
                {
                    if (_table.GetFirstChild<W.TableProperties>() == null)
                        _table.AddChild(new W.TableProperties());
                    var tblPr = _table.GetFirstChild<W.TableProperties>();
                    if (tblPr.TableLook == null)
                        tblPr.TableLook = new W.TableLook();
                    tblPr.TableLook.FirstColumn = value.Val;
                }
            }
        }

        /// <summary>
        /// Specifies that the last column format shall be applied to the table.
        /// </summary>
        public BooleanValue LastColumnEnabled
        {
            get
            {
                W.TableLook look = null;
                if (_table != null)
                {
                    look = _table.GetFirstChild<W.TableProperties>()?.TableLook;
                }
                if (look?.LastColumn != null) return look.LastColumn.Value;
                return null;
            }
            set
            {
                if (_table != null)
                {
                    if (_table.GetFirstChild<W.TableProperties>() == null)
                        _table.AddChild(new W.TableProperties());
                    var tblPr = _table.GetFirstChild<W.TableProperties>();
                    if (tblPr.TableLook == null)
                        tblPr.TableLook = new W.TableLook();
                    tblPr.TableLook.LastColumn = value.Val;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the table is floating.
        /// </summary>
        public BooleanValue WrapTextAround
        {
            get
            {
                if(_table != null)
                {
                    var tblPr = _table.GetFirstChild<W.TableProperties>();
                    return tblPr?.TablePositionProperties != null;
                }
                return null;
            }
            set
            {
                if(_table != null)
                {
                    if (_table.GetFirstChild<W.TableProperties>() == null)
                        _table.AddChild(new W.TableProperties());
                    var tblPr = _table.GetFirstChild<W.TableProperties>();
                    if (value)
                    {
                        if (tblPr.TablePositionProperties == null)
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
        }
        #endregion
    }

}
