using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;
using Berry.Docx.Field;
using Berry.Docx.Formatting;

namespace Berry.Docx.Documents
{
    /// <summary>
    /// Represents a table of contents.
    /// </summary>
    public class TableOfContent
    {
        #region Private Members
        private Document _doc;
        private SdtBlock _sdt;
        private int _startLevel = 1;
        private int _endLevel = 3;
        private bool _useHyperlink = true;
        private bool _showPageNum = true;
        private bool _pageNumAlignRight = true;
        private TabStopLeader _tabStopLeader = TabStopLeader.Dot;
        #endregion

        #region Constructors
        internal TableOfContent(Document doc, SdtBlock sdt, int startLevel, int endLevel)
        {
            _doc = doc;
            _sdt = sdt;
            _startLevel = startLevel;
            _endLevel = endLevel;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// The start value of paragraph outline level. From 1 to 9.
        /// </summary>
        public int StartOutlineLevel
        {
            get => _startLevel;
            set
            {
                if (_startLevel == value) return;
                _startLevel = value;
                Reset();
            }
        }

        /// <summary>
        /// The end value of paragraph outline level. From 1 to 9.
        /// </summary>
        public int EndOutlineLevel
        {
            get => _endLevel;
            set
            {
                if(_endLevel == value) return;
                _endLevel = value;
                Reset();
            }
        }

        /// <summary>
        /// Gets or sets a value indicates whether use hyperlink in the toc paragraph.
        /// The default value is true.
        /// </summary>
        public bool UseHyperlink
        {
            get => _useHyperlink;
            set
            {
                if(_useHyperlink == value) return;
                _useHyperlink = value;
                Reset();
            }
        }

        /// <summary>
        /// Gets or sets a value indicates whether show page number at the end of the toc paragraph.
        /// The default value is true.
        /// </summary>
        public bool ShowPageNumber
        {
            get => _showPageNum;
            set
            {
                if(_showPageNum == value) return;
                _showPageNum = value;
                Reset();
            }
        }

        /// <summary>
        /// Gets or sets a value indicates whether align right the page number if the page number is shown.
        /// The default value is true.
        /// </summary>
        public bool PageNumberAlignRight
        {
            get => _pageNumAlignRight;
            set
            {
                if (_pageNumAlignRight == value) return;
                _pageNumAlignRight = value;
                Reset();
            }
        }

        /// <summary>
        /// Gets or sets the tab leader character between content and page number.
        /// </summary>
        public TabStopLeader TabStopLeader
        {
            get => _tabStopLeader;
            set => _tabStopLeader = value;
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Update the content of the toc. This method does not guarantee correct results
        /// and does not update page numbers.
        /// </summary>
        public void Update()
        {
            if (_startLevel < 1 || _startLevel > 9
                || _endLevel < 1 || _endLevel > 9
                || _startLevel > _endLevel)
            {
                throw new ArgumentOutOfRangeException("The value of the outline level is invalid. " +
                    "And the start value can not greater than the end value.");
            }
            Reset();
            float pos = 415;
            var section = _doc.Sections.Where(s => s.ChildObjects.Contains(_sdt)).FirstOrDefault();
            if(section != null)
            { 
                pos = section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin;
            }
            bool firstParagraph = false;
            foreach(var p in _doc.Paragraphs)
            {
                OutlineLevelType outlineLevel = p.Format.OutlineLevel;
                if (string.IsNullOrEmpty(p.Text)
                    || outlineLevel == OutlineLevelType.BodyText 
                    || (int)outlineLevel + 1 < _startLevel 
                    || (int)outlineLevel + 1 > _endLevel) continue;
                List<ParagraphItem> paragraphItems = new List<ParagraphItem>();
                List <TabStop> tabStops = new List<TabStop>();
                // Add bookmark
                string bookmarkId = Bookmark.CreateNewId(_doc);
                string bookmarkName = $"_Toc{bookmarkId}";
                BookmarkStart bookmarkStart = new BookmarkStart(_doc, bookmarkId, bookmarkName);
                BookmarkEnd bookmarkEnd = new BookmarkEnd(_doc, bookmarkId);
                p.ChildItems.InsertAt(bookmarkStart, 0);
                p.ChildItems.Add(bookmarkEnd);
                
                // Add ListText
                if (!string.IsNullOrEmpty(p.ListText))
                {
                    paragraphItems.Add(new TextRange(_doc, p.ListText));
                }
                if (p.ListFormat.CurrentLevel != null)
                {
                    var level = p.ListFormat.CurrentLevel;
                    // Add Tab
                    if (level.SuffixCharacter == LevelSuffixCharacter.Tab)
                    {
                        paragraphItems.Add(new Tab(_doc));
                        tabStops.Add(new TabStop(22, TabStopStyle.Left, TabStopLeader.None));
                    }
                    // Add Space
                    else if (level.SuffixCharacter == LevelSuffixCharacter.Space)
                    {
                        paragraphItems.Add(new TextRange(_doc, " "));
                    }
                }
                // Add Text
                paragraphItems.Add(new TextRange(_doc, p.Text));
                // Add Page number
                if (_showPageNum)
                {
                    if (_pageNumAlignRight)
                    {
                        // Add Tab
                        paragraphItems.Add(new Tab(_doc));
                        tabStops.Add(new TabStop(pos, TabStopStyle.Right, _tabStopLeader));
                    }
                    else
                    {
                        paragraphItems.Add(new TextRange(_doc, " "));
                    }
                    // Add PAGEREF Field
                    var fieldBegin = new FieldChar(_doc, FieldCharType.Begin);
                    var fieldCode = new FieldCode(_doc, $" PAGEREF {bookmarkName} \\h ");
                    var fieldSeparate = new FieldChar(_doc, FieldCharType.Separate);
                    var text = new TextRange(_doc, "1");
                    var fieldEnd = new FieldChar(_doc, FieldCharType.End);
                    paragraphItems.Add(fieldBegin);
                    paragraphItems.Add(fieldCode);
                    paragraphItems.Add(fieldSeparate);
                    paragraphItems.Add(text);
                    paragraphItems.Add(fieldEnd);
                }
                
                Paragraph paragraph = null;
                if (!firstParagraph)
                {
                    paragraph = _sdt.Content.ChildObjects[1] as Paragraph;
                    firstParagraph = true;
                }
                else
                {
                    paragraph = new Paragraph(_doc);
                    _sdt.Content.ChildObjects.Last().InsertBeforeSelf(paragraph);
                }
                foreach (var tabStop in tabStops) paragraph.Format.Tabs.Add(tabStop);
                if (_useHyperlink)
                {
                    Hyperlink hyperlink = new Hyperlink(_doc, HyperlinkTargetType.Bookmark, bookmarkName, "");
                    foreach (var item in paragraphItems) hyperlink.ChildObjects.Add(item);
                    paragraph.ChildItems.Add(hyperlink);
                }
                else
                {
                    foreach (var item in paragraphItems) paragraph.ChildItems.Add(item);
                }
                
                // Apply style
                switch (outlineLevel)
                {
                    case OutlineLevelType.Level1:
                        paragraph.ApplyStyle(BuiltInStyle.TOC1);
                        break;
                    case OutlineLevelType.Level2:
                        paragraph.ApplyStyle(BuiltInStyle.TOC2);
                        break;
                    case OutlineLevelType.Level3:
                        paragraph.ApplyStyle(BuiltInStyle.TOC3);
                        break;
                    case OutlineLevelType.Level4:
                        paragraph.ApplyStyle(BuiltInStyle.TOC4);
                        break;
                    case OutlineLevelType.Level5:
                        paragraph.ApplyStyle(BuiltInStyle.TOC5);
                        break;
                    case OutlineLevelType.Level6:
                        paragraph.ApplyStyle(BuiltInStyle.TOC6);
                        break;
                    case OutlineLevelType.Level7:
                        paragraph.ApplyStyle(BuiltInStyle.TOC7);
                        break;
                    case OutlineLevelType.Level8:
                        paragraph.ApplyStyle(BuiltInStyle.TOC8);
                        break;
                    case OutlineLevelType.Level9:
                        paragraph.ApplyStyle(BuiltInStyle.TOC9);
                        break;
                    default:
                        break;
                }
            }
        }
        #endregion

        #region Private Methods
        private void Reset()
        {
            _sdt.Content.ChildObjects.Clear();

            Paragraph tocP = new Paragraph(_doc);
            tocP.AppendText("目录");
            tocP.ApplyStyle(BuiltInStyle.TOCHeading);
            tocP.ListFormat.ClearFormatting();

            Paragraph tocBegin = new Paragraph(_doc);
            Paragraph tocEnd = new Paragraph(_doc);
            var fieldBegin = new FieldChar(_doc, FieldCharType.Begin);
            // 生成域代码
            StringBuilder code = new StringBuilder();
            // 大纲级别
            code.Append($" TOC \\o \"{_startLevel}-{_endLevel}\" ");
            // 不显示页码
            if (!_showPageNum) code.Append("\\n ");
            // 显示页码时页码不右对齐
            if (_showPageNum && !_pageNumAlignRight) code.Append("\\p \" \" ");
            // 使用超链接
            if (_useHyperlink) code.Append("\\h ");
            code.Append("\\z \\u ");

            var fieldCode = new FieldCode(_doc, code.ToString());
            var fieldSeparate = new FieldChar(_doc, FieldCharType.Separate);
            var fieldEnd = new FieldChar(_doc, FieldCharType.End);
            tocBegin.ChildItems.Add(fieldBegin);
            tocBegin.ChildItems.Add(fieldCode);
            tocBegin.ChildItems.Add(fieldSeparate);
            tocEnd.ChildItems.Add(fieldEnd);

            _sdt.Content.ChildObjects.Add(tocP);
            _sdt.Content.ChildObjects.Add(tocBegin);
            _sdt.Content.ChildObjects.Add(tocEnd);
        }
        #endregion
    }
}
