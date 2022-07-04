using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the numbering format.
    /// </summary>
    public class ListFormat
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Paragraph _ownerParagraph;
        private readonly W.Style _ownerStyle;
        #endregion

        #region Constructors
        internal ListFormat(Document doc, W.Paragraph ownerParagraph)
        {
            _doc = doc;
            _ownerParagraph = ownerParagraph;
        }

        internal ListFormat(Document doc, W.Style ownerStyle)
        {
            _doc = doc;
            _ownerStyle = ownerStyle;
        }

        #endregion

        #region Public Properties

        public ListStyle CurrentStyle
        {
            get
            {
                if(_ownerParagraph != null) // Paragraph
                {
                    W.NumberingProperties numPr = _ownerParagraph.ParagraphProperties?.NumberingProperties;
                    if (numPr?.NumberingId != null && numPr?.NumberingLevelReference != null)
                    {
                        W.AbstractNum abstractNum = GetAbstractNumByID(_doc, numPr.NumberingId.Val);
                        if(abstractNum != null) return new ListStyle(_doc, abstractNum);
                    }
                    else
                    {
                        W.AbstractNum abstractNum = GetStyleAbstractNumRecursively(_doc, _ownerParagraph.GetStyle(_doc));
                        if (abstractNum != null) return new ListStyle(_doc, abstractNum);
                    }
                }
                else // Style
                {
                    W.AbstractNum abstractNum = GetStyleAbstractNumRecursively(_doc, _ownerStyle);
                    if (abstractNum != null) return new ListStyle(_doc, abstractNum);
                }
                return null;
            }
        }

        public ListLevel CurrentLevel
        {
            get
            {
                if (_ownerParagraph != null) // Paragraph
                {
                    W.NumberingProperties numPr = _ownerParagraph.ParagraphProperties?.NumberingProperties;
                    if (numPr?.NumberingId != null && numPr?.NumberingLevelReference != null)
                    {
                        W.AbstractNum abstractNum = GetAbstractNumByID(_doc, numPr.NumberingId.Val);
                        if (abstractNum != null)
                        {
                            W.Level level = GetAbstractNumLevel(abstractNum, numPr.NumberingLevelReference.Val.Value);
                            if(level != null) return new ListLevel(_doc, abstractNum, level);
                        }
                    }
                    else
                    {
                        W.Style style = _ownerParagraph.GetStyle(_doc);
                        W.AbstractNum abstractNum = GetStyleAbstractNumRecursively(_doc, style);
                        if (abstractNum != null)
                        {
                            W.Level level = GetAbstractNumLevel(abstractNum, style.StyleId);
                            if (level != null) return new ListLevel(_doc, abstractNum, level);
                        }
                    }
                }
                else // Style
                {
                    W.AbstractNum abstractNum = GetStyleAbstractNumRecursively(_doc, _ownerStyle);
                    if (abstractNum != null)
                    {
                        W.Level level = GetAbstractNumLevel(abstractNum, _ownerStyle.StyleId);
                        if (level != null) return new ListLevel(_doc, abstractNum, level);
                    }
                }
                return null;
            }
        }

        public int ListLevelNumber
        {
            get
            {
                return CurrentLevel?.ListLevelNumber ?? 0;
            }
            set
            {
                InitNumberingProperties();
                if (_ownerParagraph != null)
                {
                    W.NumberingProperties num = _ownerParagraph.ParagraphProperties.NumberingProperties;
                    num.NumberingLevelReference = new W.NumberingLevelReference() { Val = value - 1 };
                }
                else
                {
                    if (CurrentStyle == null) return;
                    foreach (ListLevel level in CurrentStyle.Levels)
                    {
                        if (level.ParagraphStyleId == _ownerStyle.StyleId)
                            level.ParagraphStyleId = null;
                    }
                    CurrentStyle.Levels[value - 1].ParagraphStyleId = _ownerStyle.StyleId;
                }
            }
        }
        #endregion

        public void ApplyStyle(ListStyle style, int levelNumber)
        {
            InitNumberingProperties();
            if (_ownerParagraph != null)
            {
                W.NumberingProperties num = _ownerParagraph.ParagraphProperties.NumberingProperties;
                num.NumberingId = new W.NumberingId() { Val = style.NumberID };
                num.NumberingLevelReference = new W.NumberingLevelReference() { Val = levelNumber - 1 };
            }
            else
            {
                W.NumberingProperties num = _ownerStyle.StyleParagraphProperties.NumberingProperties;
                num.NumberingId = new W.NumberingId() { Val = style.NumberID };
                foreach(ListLevel level in style.Levels)
                {
                    if (level.ParagraphStyleId == _ownerStyle.StyleId)
                        level.ParagraphStyleId = null;
                }
                style.Levels[levelNumber - 1].ParagraphStyleId = _ownerStyle.StyleId;
            }
        }

        public void ApplyStyle(string styleName, int levelNumber)
        {
            ListStyle style = _doc.ListStyles.FindByName(styleName);
            if (style == null) return;
            ApplyStyle(style, levelNumber);
        }


        private void InitNumberingProperties()
        {
            if (_ownerParagraph != null)
            {
                if (_ownerParagraph.ParagraphProperties == null)
                    _ownerParagraph.ParagraphProperties = new W.ParagraphProperties();
                if (_ownerParagraph.ParagraphProperties.NumberingProperties == null)
                    _ownerParagraph.ParagraphProperties.NumberingProperties = new W.NumberingProperties();
            }
            else
            {
                if (_ownerStyle.StyleParagraphProperties == null)
                    _ownerStyle.StyleParagraphProperties = new W.StyleParagraphProperties();
                if (_ownerStyle.StyleParagraphProperties.NumberingProperties == null)
                    _ownerStyle.StyleParagraphProperties.NumberingProperties = new W.NumberingProperties();
            }
        }
        private static W.AbstractNum GetAbstractNumByID(Document doc, int numId)
        {
            W.Numbering numbering = doc.Package.MainDocumentPart.NumberingDefinitionsPart?.Numbering;
            if (numbering == null) return null;
            W.NumberingInstance num = numbering.Elements<W.NumberingInstance>().Where(n => n.NumberID == numId).FirstOrDefault();
            if (num == null) return null;
            int abstractNumId = num.AbstractNumId.Val;
            W.AbstractNum abstractNum = numbering.Elements<W.AbstractNum>().Where(a => a.AbstractNumberId == abstractNumId).FirstOrDefault();
            return abstractNum;
        }

        private static W.Level GetAbstractNumLevel(W.AbstractNum num, int levelIndex)
        {
            return num?.Elements<W.Level>().Where(l => l.LevelIndex == levelIndex).FirstOrDefault();
        }

        private static W.Level GetAbstractNumLevel(W.AbstractNum num, string styleId)
        {
            return num?.Elements<W.Level>().Where(l => l.ParagraphStyleIdInLevel?.Val == styleId).FirstOrDefault();
        }

        private static W.AbstractNum GetStyleAbstractNumRecursively(Document doc, W.Style style)
        {
            if (style.StyleParagraphProperties?.NumberingProperties?.NumberingId != null)
            {
                int numId = style.StyleParagraphProperties.NumberingProperties.NumberingId.Val;
                W.AbstractNum abstractNum = GetAbstractNumByID(doc, numId);
                if(abstractNum != null) return abstractNum;
            }
            W.Style baseStyle = style.GetBaseStyle(doc);
            if (baseStyle != null)
            {
                return GetStyleAbstractNumRecursively(doc, baseStyle);
            }
            return null;
        }


    }
}
