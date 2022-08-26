using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the style of one region in the table. Look <see cref="TableStyle"/> for supporting regions.
    /// <para>表示表格中某一区域的样式. <see cref="TableStyle"/> 中定义了支持的表格区域.</para>
    /// </summary>
    public class TableRegionStyle
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Style _style;
        private readonly TableRegionType _region;
        private readonly CharacterFormat _cFormat;
        private readonly ParagraphFormat _pFormat;
        private readonly TableBorders _borders;
        #endregion

        #region Constructors
        internal TableRegionStyle(Document doc, W.Style style, TableRegionType region)
        {
            _doc = doc;
            _style = style;
            _region = region;
            _cFormat = new CharacterFormat(doc, style, region);
            _pFormat = new ParagraphFormat(doc, style, region);
            _borders = new TableBorders(doc, style, region);
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the table cell character format.
        /// </summary>
        public CharacterFormat CharacterFormat => _cFormat;

        /// <summary>
        /// Gets the table cell paragraph format.
        /// </summary>
        public ParagraphFormat ParagraphFormat => _pFormat;

        /// <summary>
        /// Gets or sets the horizontal alignment.
        /// </summary>
        public TableRowAlignment HorizontalAlignment
        {
            get
            {
                W.TableJustification jc = null;
                if(_region == TableRegionType.WholeTable)
                {
                    jc = _style.StyleTableProperties?.TableJustification;
                    if(jc == null)
                    {
                        W.Style baseStyle = _style.GetBaseStyle(_doc);
                        if (baseStyle != null)
                        {
                            return new TableRegionStyle(_doc, baseStyle, _region).HorizontalAlignment;
                        }
                        return TableRowAlignment.Left;
                    }
                    return jc.Val.Value.Convert<TableRowAlignment>();
                }
                else
                {
                    W.TableStyleOverrideValues type = _region.Convert<W.TableStyleOverrideValues>();
                    W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                    jc = tblStylePr?.TableStyleConditionalFormattingTableProperties?.TableJustification;
                    if (jc == null)
                    {
                        return new TableRegionStyle(_doc, _style, TableRegionType.WholeTable).HorizontalAlignment;
                    }
                    return jc.Val.Value.Convert<TableRowAlignment>();
                }
            }
            set
            {
                if (_region == TableRegionType.WholeTable)
                {
                    if (_style.StyleTableProperties == null)
                    {
                        _style.StyleTableProperties = new W.StyleTableProperties();
                    }
                    _style.StyleTableProperties.TableJustification = new W.TableJustification()
                    {
                        Val = value.Convert<W.TableRowAlignmentValues>()
                    };
                    if (_style.TableStyleConditionalFormattingTableRowProperties == null)
                    {
                        _style.TableStyleConditionalFormattingTableRowProperties = new W.TableStyleConditionalFormattingTableRowProperties();
                    }
                    _style.TableStyleConditionalFormattingTableRowProperties.AddChild(new W.TableJustification() { Val = value.Convert<W.TableRowAlignmentValues>() });
                }
                else if(_region == TableRegionType.FirstRow || _region == TableRegionType.LastRow)
                {
                    W.TableStyleOverrideValues type = _region.Convert<W.TableStyleOverrideValues>();
                    if (!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                    {
                        _style.Append(new W.TableStyleProperties() { Type = type });
                    }
                    W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                    if (tblStylePr.TableStyleConditionalFormattingTableProperties == null)
                    {
                        tblStylePr.TableStyleConditionalFormattingTableProperties = new W.TableStyleConditionalFormattingTableProperties();
                    }
                    if (tblStylePr.TableStyleConditionalFormattingTableProperties.TableJustification == null)
                    {
                        tblStylePr.TableStyleConditionalFormattingTableProperties.TableJustification = new W.TableJustification();
                    }
                    tblStylePr.TableStyleConditionalFormattingTableProperties.TableJustification = new W.TableJustification()
                    {
                        Val = value.Convert<W.TableRowAlignmentValues>()
                    };
                    if(tblStylePr.TableStyleConditionalFormattingTableRowProperties == null)
                    {
                        tblStylePr.TableStyleConditionalFormattingTableRowProperties = new W.TableStyleConditionalFormattingTableRowProperties();
                    }
                    tblStylePr.TableStyleConditionalFormattingTableRowProperties.AddChild(new W.TableJustification() { Val = value.Convert<W.TableRowAlignmentValues>() });
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether allow row to break across pages.
        /// </summary>
        public bool AllowBreakAcrossPages
        {
            get
            {
                W.CantSplit cantSplit = null;
                if (_region == TableRegionType.WholeTable)
                {
                    cantSplit = _style.TableStyleConditionalFormattingTableRowProperties?.GetFirstChild<W.CantSplit>();
                    if (cantSplit == null)
                    {
                        W.Style baseStyle = _style.GetBaseStyle(_doc);
                        if (baseStyle != null)
                        {
                            return new TableRegionStyle(_doc, baseStyle, _region).AllowBreakAcrossPages;
                        }
                        return true;
                    }
                    if (cantSplit.Val == null) return false;
                    return cantSplit.Val.Value == W.OnOffOnlyValues.Off;
                }
                return true;
            }
            set
            {
                if (_region != TableRegionType.WholeTable) return;
                if (value)
                {
                    _style.TableStyleConditionalFormattingTableRowProperties?.GetFirstChild<W.CantSplit>()?.Remove();
                }
                else
                {
                    if (_style.TableStyleConditionalFormattingTableRowProperties == null)
                    {
                        _style.TableStyleConditionalFormattingTableRowProperties = new W.TableStyleConditionalFormattingTableRowProperties();
                    }
                    _style.TableStyleConditionalFormattingTableRowProperties.AddChild(new W.CantSplit());
                }
            }
        }

        /// <summary>
        /// Gets or sets the table cell vertical alignment.
        /// </summary>
        public TableCellVerticalAlignment VerticalCellAlignment
        {
            get
            {
                W.TableCellVerticalAlignment vAlign = null;
                if(_region == TableRegionType.WholeTable)
                {
                    vAlign = _style.StyleTableCellProperties?.TableCellVerticalAlignment;
                    if(vAlign == null)
                    {
                        W.Style baseStyle = _style.GetBaseStyle(_doc);
                        if(baseStyle != null)
                        {
                            return new TableRegionStyle(_doc, baseStyle, _region).VerticalCellAlignment;
                        }
                        return TableCellVerticalAlignment.Top;
                    }
                    return vAlign.Val.Value.Convert<TableCellVerticalAlignment>();
                }
                else
                {
                    W.TableStyleOverrideValues type = _region.Convert<W.TableStyleOverrideValues>();
                    W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                    vAlign = tblStylePr?.TableStyleConditionalFormattingTableCellProperties?.TableCellVerticalAlignment;
                    if(vAlign == null)
                    {
                        return new TableRegionStyle(_doc, _style, TableRegionType.WholeTable).VerticalCellAlignment;
                    }
                    return vAlign.Val.Value.Convert<TableCellVerticalAlignment>();
                }
            }
            set
            {
                if (_region == TableRegionType.WholeTable)
                {
                    if (_style.StyleTableCellProperties == null)
                        _style.StyleTableCellProperties = new W.StyleTableCellProperties();
                    if(value == TableCellVerticalAlignment.Top)
                    {
                        W.Style baseStyle = _style.GetBaseStyle(_doc);
                        if (baseStyle == null || new TableRegionStyle(_doc, baseStyle, _region).VerticalCellAlignment == TableCellVerticalAlignment.Top)
                        {
                            _style.StyleTableCellProperties.TableCellVerticalAlignment = null;
                            return;
                        }
                    }
                    _style.StyleTableCellProperties.TableCellVerticalAlignment = new W.TableCellVerticalAlignment()
                    {
                        Val = value.Convert<W.TableVerticalAlignmentValues>()
                    };
                }
                else
                {
                    W.TableStyleOverrideValues type = _region.Convert<W.TableStyleOverrideValues>();
                    if(!_style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).Any())
                    {
                        _style.Append(new W.TableStyleProperties() { Type = type });
                    }
                    W.TableStyleProperties tblStylePr = _style.Elements<W.TableStyleProperties>().Where(t => t.Type == type).FirstOrDefault();
                    if (tblStylePr.TableStyleConditionalFormattingTableCellProperties == null)
                    {
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties = new W.TableStyleConditionalFormattingTableCellProperties();
                    }
                    if (tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellVerticalAlignment == null)
                    {
                        tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellVerticalAlignment = new W.TableCellVerticalAlignment();
                    }
                    tblStylePr.TableStyleConditionalFormattingTableCellProperties.TableCellVerticalAlignment = new W.TableCellVerticalAlignment()
                    {
                        Val = value.Convert<W.TableVerticalAlignmentValues>()
                    };
                }
            }
        }

        /// <summary>
        /// Gets the table cell borders.
        /// </summary>
        public TableBorders Borders => _borders;
        #endregion
    }

    internal enum TableRegionType
    {
        WholeTable = 0,
        FirstRow = 1,
        LastRow = 2,
        FirstColumn = 3,
        LastColumn = 4
    }
}
