using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OOxml = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// 字符格式
    /// </summary>
    public class CharacterFormat
    {
        private Document _doc;

        #region TextRange
        private OOxml.Run _ownerRun = null;
        private RunPropertiesHolder _curRHld = null;
        private CharacterFormat _inheritFromParagraphFormat = null;
        #endregion

        #region Paragraph
        private OOxml.Paragraph _ownerParagraph = null;
        private RunPropertiesHolder _curPHld = null;
        private CharacterFormat _inheritFromStyleFormat = null;
        #endregion

        #region Style
        private OOxml.Style _ownerStyle = null;
        private RunPropertiesHolder _curSHld = null;
        private CharacterFormat _inheritFromBaseStyleFormat = null;
        #endregion

        #region Formats
        private string _fontCN = "宋体";
        private string _fontEN = "Times New Roman";
        private float _fontSize = 10.5F;
        private float _fontSizeCs = 10.5F;
        private bool _bold = false;
        private bool _italic = false;
        #endregion

        public CharacterFormat() { }

        public CharacterFormat(Document doc, OOxml.Run ownerRun)
        {
            _doc = doc;
            _ownerRun = ownerRun;
            if (ownerRun.RunProperties == null)
                ownerRun.RunProperties = new OOxml.RunProperties();
            _curRHld = new RunPropertiesHolder(doc.Package, ownerRun.RunProperties);
            if(ownerRun.Parent != null)
                _inheritFromParagraphFormat = new CharacterFormat(doc, ownerRun.Parent as OOxml.Paragraph);
        }

        public CharacterFormat(Document doc, OOxml.Paragraph ownerParagraph)
        {
            _doc = doc;
            _ownerParagraph = ownerParagraph;
            if (ownerParagraph.ParagraphProperties == null)
                ownerParagraph.ParagraphProperties = new OOxml.ParagraphProperties();
            if (ownerParagraph.ParagraphProperties.ParagraphMarkRunProperties == null)
                ownerParagraph.ParagraphProperties.ParagraphMarkRunProperties = new OOxml.ParagraphMarkRunProperties();
            _curPHld = new RunPropertiesHolder(doc.Package, ownerParagraph.ParagraphProperties.ParagraphMarkRunProperties);
            _inheritFromStyleFormat = new CharacterFormat(doc, ownerParagraph.GetStyle(doc));
        }
        
        public CharacterFormat(Document doc, OOxml.Style ownerStyle)
        {
            _doc = doc;
            _ownerStyle = ownerStyle;
            if (ownerStyle.StyleRunProperties == null)
                ownerStyle.StyleRunProperties = new OOxml.StyleRunProperties();
            _curSHld = new RunPropertiesHolder(doc.Package, ownerStyle.StyleRunProperties);
            _inheritFromBaseStyleFormat = GetStyleCharacterFormatRecursively(ownerStyle);
        }

        private CharacterFormat GetStyleCharacterFormatRecursively(OOxml.Style style)
        {
            CharacterFormat format = new CharacterFormat();
            CharacterFormat baseFormat = new CharacterFormat();
            // 获取默认值
            OOxml.Styles styles = style.Parent as OOxml.Styles;
            if (styles.DocDefaults != null && styles.DocDefaults.RunPropertiesDefault != null
                && styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle != null)
            {
                RunPropertiesHolder rPr = new RunPropertiesHolder(_doc.Package, styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle);
                baseFormat.FontCN = rPr.FontCN;
                baseFormat.FontEN = rPr.FontEN;
                baseFormat.FontSize = rPr.FontSize;
                baseFormat.FontSizeCs = rPr.FontSizeCs;
                baseFormat.Bold = rPr.Bold ?? false;
                baseFormat.Italic = rPr.Italic ?? false;
            }
            //获取基类样式
            OOxml.Style baseStyle = style.GetBaseStyle();
            if (baseStyle != null)
                baseFormat = GetStyleCharacterFormatRecursively(baseStyle);
            if (style.StyleRunProperties == null) style.StyleRunProperties = new OOxml.StyleRunProperties();
            RunPropertiesHolder curSHld = new RunPropertiesHolder(_doc.Package, style.StyleRunProperties);

            format.FontCN = curSHld.FontCN ?? baseFormat.FontCN;
            format.FontEN = curSHld.FontEN ?? baseFormat.FontEN;
            format.FontSize = curSHld.FontSize > 0 ? curSHld.FontSize : baseFormat.FontSize;
            format.FontSizeCs = curSHld.FontSizeCs > 0 ? curSHld.FontSizeCs : baseFormat.FontSizeCs;
            format.Bold = curSHld.Bold ?? baseFormat.Bold;
            format.Italic = curSHld.Italic ?? baseFormat.Italic;
            return format;
        }

        /// <summary>
        /// 中文字体
        /// </summary>
        public string FontCN
        {
            get
            {
                if(_ownerRun != null)
                {
                    InitRun();
                    return _curRHld.FontCN ?? (_inheritFromParagraphFormat != null ? _inheritFromParagraphFormat.FontCN : string.Empty);
                }
                else if(_ownerParagraph != null)
                {
                    return _curPHld.FontCN ?? _inheritFromStyleFormat.FontCN;
                }
                else if(_ownerStyle != null)
                {
                    return _curSHld.FontCN ?? _inheritFromBaseStyleFormat.FontCN;
                }
                else
                {
                    return _fontCN;
                }
            }
            set
            {
                if (_ownerRun != null)
                {
                    _curRHld.FontCN = value;
                }
                else if(_ownerParagraph != null)
                {
                    _curPHld.FontCN = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.FontCN = value;
                }
                else
                {
                    _fontCN = value;
                }
            }
        }

        /// <summary>
        /// 英文字体
        /// </summary>
        public string FontEN
        {
            get
            {
                if (_ownerRun != null)
                {
                    InitRun();
                    return _curRHld.FontEN ?? (_inheritFromParagraphFormat != null ? _inheritFromParagraphFormat.FontEN : string.Empty);
                }
                else if (_ownerParagraph != null)
                {
                    return _curPHld.FontEN ?? _inheritFromStyleFormat.FontEN;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.FontEN ?? _inheritFromBaseStyleFormat.FontEN;
                }
                else
                {
                    return _fontEN;
                }
            }
            set
            {
                if (_ownerRun != null)
                {
                    _curRHld.FontEN = value;
                }
                else if (_ownerParagraph != null)
                {
                    _curPHld.FontEN = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.FontEN = value;
                }
                else
                {
                    _fontEN = value;
                }
            }
        }
        /// <summary>
        /// 字号
        /// </summary>
        public float FontSize
        {
            get
            {
                if (_ownerRun != null)
                {
                    InitRun();
                    if(_curRHld.FontSize > 0)
                    {
                        return _curRHld.FontSize;
                    }
                    return _inheritFromParagraphFormat != null ? _inheritFromParagraphFormat.FontSize : 0;
                }
                else if (_ownerParagraph != null)
                {
                    return _curPHld.FontSize > 0 ? _curPHld.FontSize : _inheritFromStyleFormat.FontSize;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.FontSize > 0 ? _curSHld.FontSize : _inheritFromBaseStyleFormat.FontSize;
                }
                else
                {
                    return _fontSize;
                }
            }
            set
            {
                if(_ownerRun != null)
                {
                    _curRHld.FontSize = value;
                }
                else if (_ownerParagraph != null)
                {
                    _curPHld.FontSize = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.FontSize = value;
                }
                else
                {
                    _fontSize = value;
                }
            }
        }
        /// <summary>
        /// 中文字号
        /// </summary>
        public string FontSizeCN
        {
            get
            {
                Dictionary<float, string> sizeList = new Dictionary<float, string> { 
                    { 6.5F, "小六" }, { 7.5F, "六号" }, { 9, "小五" }, { 10.5F, "五号" }, { 12, "小四" }, 
                    { 14, "四号" }, { 15, "小三" }, { 16, "三号" }, { 18, "小二" }, { 22, "二号" }, 
                    { 24, "小一" }, { 26, "一号" }, { 36, "小初" }, { 42, "初号" } 
                };
                float size = FontSize;
                return sizeList.ContainsKey(size) ? sizeList[size] : size.ToString();
            }
            set
            {
                Dictionary<string, float> sizeList = new Dictionary<string, float> {
                    { "小六", 6.5F }, { "六号", 7.5F }, { "小五", 9 }, { "五号", 10.5F }, { "小四", 12 },
                    { "四号", 14 }, { "小三", 15 }, { "三号", 16 }, { "小二", 18 }, { "二号", 22 },
                    { "小一", 24 }, { "一号", 26 }, { "小初", 36 }, { "初号", 42 }
                };
                if (sizeList.ContainsKey(value))
                    FontSize = sizeList[value];
            }
        }
        /// <summary>
        /// 字号
        /// </summary>
        public float FontSizeCs
        {
            get
            {
                if(_ownerRun != null)
                {
                    InitRun();
                    if (_curRHld.FontSizeCs > 0)
                    {
                        return _curRHld.FontSizeCs;
                    }
                    return _inheritFromParagraphFormat != null ? _inheritFromParagraphFormat.FontSizeCs : 0;
                }
                else if (_ownerParagraph != null)
                {
                    return _curPHld.FontSizeCs > 0 ? _curPHld.FontSizeCs : _inheritFromStyleFormat.FontSizeCs;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.FontSizeCs > 0 ?_curSHld.FontSizeCs : _inheritFromBaseStyleFormat.FontSizeCs;
                }
                else
                {
                    return _fontSizeCs;
                }
            }
            set
            {
                if (_ownerRun != null)
                {
                    _curRHld.FontSizeCs = value;
                }
                else if (_ownerParagraph != null)
                {
                    _curPHld.FontSizeCs = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.FontSizeCs = value;
                }
                else
                {
                    _fontSizeCs = value;
                }
            }
        }

        /// <summary>
        /// 加粗
        /// </summary>
        public bool Bold
        {
            get
            {
                if(_ownerRun != null)
                {
                    return _curRHld.Bold ?? (_inheritFromParagraphFormat != null ? _inheritFromParagraphFormat.Bold : false);
                }
                else if (_ownerParagraph != null)
                {
                    return _curPHld.Bold ?? _inheritFromStyleFormat.Bold;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.Bold ?? _inheritFromBaseStyleFormat.Bold;
                }
                else
                {
                    return _bold;
                }
            }
            set
            {
                if(_ownerRun != null)
                {
                    _curRHld.Bold = value;
                }
                else if (_ownerParagraph != null)
                {
                    _curPHld.Bold = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.Bold = value;
                }
                else
                {
                    _bold = value;
                }
            }
        }
        /// <summary>
        /// 斜体
        /// </summary>
        public bool Italic
        {
            get
            {
                if(_ownerRun != null)
                {
                    return _curRHld.Italic ?? (_inheritFromParagraphFormat != null ? _inheritFromParagraphFormat.Italic : false);
                }
                if (_ownerParagraph != null)
                {
                    return _curPHld.Italic ?? _inheritFromStyleFormat.Italic;
                }
                else if (_ownerStyle != null)
                {
                    return _curSHld.Italic ?? _inheritFromBaseStyleFormat.Italic;
                }
                else
                {
                    return _italic;
                }
            }
            set
            {
                if (_ownerRun != null)
                {
                    _curRHld.Italic = value;
                }
                else if (_ownerParagraph != null)
                {
                    _curPHld.Italic = value;
                }
                else if (_ownerStyle != null)
                {
                    _curSHld.Italic = value;
                }
                else
                {
                    _italic = value;
                }
            }
        }

        private void InitRun()
        {
            if(_ownerRun.Parent != null && _inheritFromParagraphFormat == null)
                _inheritFromParagraphFormat = new CharacterFormat(_doc, _ownerRun.Parent as OOxml.Paragraph);
        }

    }

}
