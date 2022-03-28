using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OOxml = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Documents
{
    public class Settings
    {
        private OOxml.Settings _settings;
        public Settings(OOxml.Settings settings)
        {
            _settings = settings;
        }
        internal bool EvenAndOddHeaders
        {
            get
            {
                return _settings.Elements<OOxml.EvenAndOddHeaders>().Any();
            }
            set
            {
                if (value)
                {
                    if (!_settings.Elements<OOxml.EvenAndOddHeaders>().Any())
                        _settings.AddChild(new OOxml.EvenAndOddHeaders());
                }
                else
                {
                    _settings.RemoveAllChildren<OOxml.EvenAndOddHeaders>();
                }
            }
        }

        /// <summary>
        /// 装订线位置为上，返回True，否则返回False
        /// </summary>
        internal bool GutterAtTop
        {
            get
            {
                if(_settings.GutterAtTop == null) return false;
                if (_settings.GutterAtTop.Val == null) return true;
                return _settings.GutterAtTop.Val;
            }
            set
            {
                if (value)
                    _settings.GutterAtTop = new OOxml.GutterAtTop();
                else
                    _settings.GutterAtTop = null;
            }
        }
        /// <summary>
        /// 页码范围-多页
        /// </summary>
        public MultiPage MultiPage
        {
            get
            {
                if (MirrorMargins)
                    return MultiPage.MirrorMargins;
                else if (PrintTwoOnOne)
                    return MultiPage.PrintTwoOnOne;
                else
                    return MultiPage.Normal;
            }
            set
            {
                switch (value)
                {
                    case MultiPage.MirrorMargins:
                        MirrorMargins = true;
                        PrintTwoOnOne = false;
                        break;
                    case MultiPage.PrintTwoOnOne:
                        MirrorMargins = false;
                        PrintTwoOnOne = true;
                        break;
                    default:
                        MirrorMargins = false;
                        PrintTwoOnOne = false;
                        break;
                }
            }
        }
        /// <summary>
        /// 对称页边距
        /// </summary>
        private bool MirrorMargins
        {
            get
            {
                return _settings.MirrorMargins != null;
            }
            set
            {
                if (value)
                    _settings.MirrorMargins = new OOxml.MirrorMargins();
                else
                    _settings.MirrorMargins = null;
            }
        }
        /// <summary>
        /// 拼页
        /// </summary>
        private bool PrintTwoOnOne
        {
            get
            {
                return _settings.Elements<OOxml.PrintTwoOnOne>().Count() > 0;
            }
            set
            {
                if (value)
                {
                    if (_settings.Elements<OOxml.PrintTwoOnOne>().Count() == 0)
                        _settings.AddChild(new OOxml.PrintTwoOnOne());
                }
                else
                {
                    OOxml.PrintTwoOnOne printTwoOnOne = _settings.Elements<OOxml.PrintTwoOnOne>().FirstOrDefault();
                    if (printTwoOnOne != null)
                        printTwoOnOne.Remove();
                }
            }
        }

    }
}
