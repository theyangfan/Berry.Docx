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
        /// The start value of outlinelevel.
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
        /// The end value of outlinelevel.
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
        #endregion

        #region Public Methods
        /// <summary>
        /// 
        /// </summary>
        public void Update()
        {
            int index = 1;
            foreach(var p in _doc.Paragraphs)
            {
                OutlineLevelType outlineLevel = p.Format.OutlineLevel;
                if (outlineLevel == OutlineLevelType.BodyText) continue;
                if ((int)outlineLevel + 1 < _startLevel || (int)outlineLevel + 1 > _endLevel) continue;
                List <TabStop> tabStops = new List<TabStop>();
                // Add bookmark
                string bookmarkId = Bookmark.CreateNewId(_doc);
                string bookmarkName = $"_Toc{bookmarkId}";
                BookmarkStart bookmarkStart = new BookmarkStart(_doc, bookmarkId, bookmarkName);
                BookmarkEnd bookmarkEnd = new BookmarkEnd(_doc, bookmarkId);
                p.ChildItems.InsertAt(bookmarkStart, 0);
                p.ChildItems.Add(bookmarkEnd);

                // Insert hyperlink
                Hyperlink hyperlink = new Hyperlink(_doc, HyperlinkTargetType.Bookmark, bookmarkName, "");
                if (!string.IsNullOrEmpty(p.ListText))
                {
                    hyperlink.ChildObjects.Add(new TextRange(_doc, p.ListText));
                }
                if (p.ListFormat.CurrentLevel != null)
                {
                    var level = p.ListFormat.CurrentLevel;
                    if (level.SuffixCharacter == LevelSuffixCharacter.Tab)
                    {
                        hyperlink.ChildObjects.Add(new Tab(_doc));
                        tabStops.Add(new TabStop(22, TabStopStyle.Left, TabStopLeader.None));
                    }
                    else if (level.SuffixCharacter == LevelSuffixCharacter.Space)
                    {
                        hyperlink.ChildObjects.Add(new TextRange(_doc, " "));
                    }
                }
                hyperlink.ChildObjects.Add(new TextRange(_doc, p.Text));
                hyperlink.ChildObjects.Add(new Tab(_doc));
                tabStops.Add(new TabStop(415, TabStopStyle.Right, TabStopLeader.Dot));
                var fieldBegin = new FieldChar(_doc, FieldCharType.Begin);
                var fieldCode = new FieldCode(_doc, $" PAGEREF {bookmarkName} \\h ");
                var fieldSeparate = new FieldChar(_doc, FieldCharType.Separate);
                var text = new TextRange(_doc, "1");
                var fieldEnd = new FieldChar(_doc, FieldCharType.End);
                hyperlink.ChildObjects.Add(fieldBegin);
                hyperlink.ChildObjects.Add(fieldCode);
                hyperlink.ChildObjects.Add(fieldSeparate);
                hyperlink.ChildObjects.Add(text);
                hyperlink.ChildObjects.Add(fieldEnd);

                if (outlineLevel == OutlineLevelType.Level1)
                {
                    if(index == 1)
                    {
                        var paragraph = _sdt.Content.ChildObjects[1] as Paragraph;
                        if (paragraph != null)
                        {
                            paragraph.ApplyStyle(BuiltInStyle.TOC1);
                            foreach (var tabStop in tabStops) paragraph.Format.Tabs.Add(tabStop);
                            paragraph.ChildItems.Add(hyperlink);
                        }
                    }
                    else
                    {
                        Paragraph paragraph = new Paragraph(_doc);
                        paragraph.ChildItems.Add(hyperlink);
                        paragraph.ApplyStyle(BuiltInStyle.TOC1);
                        foreach (var tabStop in tabStops) paragraph.Format.Tabs.Add(tabStop);
                        _sdt.Content.ChildObjects.Last().InsertBeforeSelf(paragraph);
                    }
                }
                if (outlineLevel == OutlineLevelType.Level2)
                {
                    if (index == 1)
                    {
                        var paragraph = _sdt.Content.ChildObjects[1] as Paragraph;
                        if (paragraph != null)
                        {
                            paragraph.ApplyStyle(BuiltInStyle.TOC2);
                            foreach (var tabStop in tabStops) paragraph.Format.Tabs.Add(tabStop);
                            paragraph.ChildItems.Add(hyperlink);
                        }
                    }
                    else
                    {
                        Paragraph paragraph = new Paragraph(_doc);
                        paragraph.ChildItems.Add(hyperlink);
                        paragraph.ApplyStyle(BuiltInStyle.TOC2);
                        foreach (var tabStop in tabStops) paragraph.Format.Tabs.Add(tabStop);
                        _sdt.Content.ChildObjects.Last().InsertBeforeSelf(paragraph);
                    }
                }
                if (outlineLevel == OutlineLevelType.Level3)
                {
                    if (index == 1)
                    {
                        var paragraph = _sdt.Content.ChildObjects[1] as Paragraph;
                        if (paragraph != null)
                        {
                            paragraph.ApplyStyle(BuiltInStyle.TOC3);
                            foreach (var tabStop in tabStops) paragraph.Format.Tabs.Add(tabStop);
                            paragraph.ChildItems.Add(hyperlink);
                        }
                    }
                    else
                    {
                        Paragraph paragraph = new Paragraph(_doc);
                        paragraph.ChildItems.Add(hyperlink);
                        paragraph.ApplyStyle(BuiltInStyle.TOC3);
                        foreach (var tabStop in tabStops) paragraph.Format.Tabs.Add(tabStop);
                        _sdt.Content.ChildObjects.Last().InsertBeforeSelf(paragraph);
                    }
                }
                ++index;
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
