using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using W = DocumentFormat.OpenXml.Wordprocessing;
using Berry.Docx.Formatting;

namespace Berry.Docx.Documents
{
    internal class Settings
    {
        private readonly Document _doc;
        private readonly W.Settings _settings;
        public Settings(Document doc, W.Settings settings)
        {
            _doc = doc;
            _settings = settings;
        }

        public W.Settings XElement => _settings;

        public bool EvenAndOddHeaders
        {
            get
            {
                return _settings.Elements<W.EvenAndOddHeaders>().Any();
            }
            set
            {
                if (value)
                {
                    if (!_settings.Elements<W.EvenAndOddHeaders>().Any())
                        _settings.AddChild(new W.EvenAndOddHeaders());
                }
                else
                {
                    _settings.RemoveAllChildren<W.EvenAndOddHeaders>();
                }
            }
        }

        /// <summary>
        /// 装订线位置为上，返回True，否则返回False
        /// </summary>
        public bool GutterAtTop
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
                    _settings.GutterAtTop = new W.GutterAtTop();
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
        public bool MirrorMargins
        {
            get
            {
                return _settings.MirrorMargins != null;
            }
            set
            {
                if (value)
                    _settings.MirrorMargins = new W.MirrorMargins();
                else
                    _settings.MirrorMargins = null;
            }
        }
        /// <summary>
        /// 拼页
        /// </summary>
        public bool PrintTwoOnOne
        {
            get
            {
                return _settings.Elements<W.PrintTwoOnOne>().Count() > 0;
            }
            set
            {
                if (value)
                {
                    if (_settings.Elements<W.PrintTwoOnOne>().Count() == 0)
                        _settings.AddChild(new W.PrintTwoOnOne());
                }
                else
                {
                    W.PrintTwoOnOne printTwoOnOne = _settings.Elements<W.PrintTwoOnOne>().FirstOrDefault();
                    if (printTwoOnOne != null)
                        printTwoOnOne.Remove();
                }
            }
        }

        public FootEndnoteFormat FootnoteFormt => new FootEndnoteFormat(_doc, this, NoteType.DocumentWideFootnote);

        public FootEndnoteFormat EndnoteFormt => new FootEndnoteFormat(_doc, this, NoteType.DocumentWideEndnote);

    }
}
