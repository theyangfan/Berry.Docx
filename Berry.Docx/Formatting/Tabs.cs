using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    public class Tabs : IEnumerable<Tab>
    {
        private readonly Document _doc;
        private readonly W.Paragraph _ownerParagraph;
        private readonly W.Style _ownerStyle;

        internal Tabs(Document doc, W.Paragraph paragraph)
        {
            _doc = doc;
            _ownerParagraph = paragraph;
        }

        internal Tabs(Document doc, W.Style style)
        {
            _doc = doc;
            _ownerStyle = style;
        }

        public void Add(float pos, TabStopStyle style, TabStopLeader leader)
        {
            Tab tab = new Tab(pos, style, leader);
            Add(tab);
        }

        public void Add(Tab tab)
        {
            if (GetTabs().Contains(tab)) return;
            if (_ownerParagraph != null)
            {
                if (_ownerParagraph.ParagraphProperties == null)
                    _ownerParagraph.ParagraphProperties = new W.ParagraphProperties();
                if (_ownerParagraph.ParagraphProperties.Tabs == null)
                    _ownerParagraph.ParagraphProperties.Tabs = new W.Tabs();
                _ownerParagraph.ParagraphProperties.Tabs.Append(tab.XETabStop);
            }
            else if (_ownerStyle != null)
            {
                if (_ownerStyle.StyleParagraphProperties == null)
                    _ownerStyle.StyleParagraphProperties = new W.StyleParagraphProperties();
                if (_ownerStyle.StyleParagraphProperties.Tabs == null)
                    _ownerStyle.StyleParagraphProperties.Tabs = new W.Tabs();
                _ownerStyle.StyleParagraphProperties.Tabs.Append(tab.XETabStop);
            }
        }

        public void Clear()
        {
            if (_ownerParagraph?.ParagraphProperties?.Tabs != null)
                _ownerParagraph.ParagraphProperties.Tabs = null;
            if (_ownerStyle != null)
            {
                if(_ownerStyle?.StyleParagraphProperties?.Tabs != null)
                    _ownerStyle.StyleParagraphProperties.Tabs = null;
                W.Style baseStyle = _ownerStyle.GetBaseStyle(_doc);
                if (baseStyle != null)
                {
                    new Tabs(_doc, baseStyle).Clear();
                }
            }
        }

        public IEnumerator<Tab> GetEnumerator()
        {
            return GetTabs().GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        private IEnumerable<Tab> GetTabs()
        {
            if (_ownerParagraph != null)
            {
                if (_ownerParagraph.ParagraphProperties?.Tabs != null)
                {
                    foreach (W.TabStop tab in _ownerParagraph.ParagraphProperties.Tabs.Elements<W.TabStop>())
                        yield return new Tab(tab);
                }
                foreach (W.TabStop tab in GetStyleTabsRecursively(_doc, _ownerParagraph.GetStyle(_doc)))
                    yield return new Tab(tab);
            }
            else if (_ownerStyle != null)
            {
                foreach (W.TabStop tab in GetStyleTabsRecursively(_doc, _ownerStyle))
                    yield return new Tab(tab);
            }
        }

        private static IEnumerable<W.TabStop> GetStyleTabsRecursively(Document doc, W.Style style)
        {
            W.Style baseStyle = style.GetBaseStyle(doc);
            if (baseStyle != null)
            {
                foreach(W.TabStop tab in GetStyleTabsRecursively(doc, baseStyle))
                {
                    yield return tab;
                }
            }
            if(style.StyleParagraphProperties?.Tabs != null)
            {
                foreach(W.TabStop tab in style.StyleParagraphProperties.Tabs.Elements<W.TabStop>())
                {
                    yield return tab;
                }
            }
        }

    }

    public class Tab : IEquatable<Tab>
    {
        private readonly W.TabStop _tab;

        public Tab() : this(new W.TabStop() { Position = 0, Val = W.TabStopValues.Clear, Leader = W.TabStopLeaderCharValues.None })
        {
        }

        public Tab(float pos, TabStopStyle style, TabStopLeader leader) : this()
        {
            Position = pos;
            Style = style;
            Leader = leader;
        }

        internal Tab(W.TabStop tab)
        {
            _tab = tab;
        }

        public float Position
        {
            get
            {
                if (_tab.Position == null) return 0;
                return _tab.Position / 20.0F;
            }
            set
            {
                _tab.Position = (int)(value * 20);
            }
        }

        public TabStopStyle Style
        {
            get
            {
                if (_tab.Val == null) return TabStopStyle.Clear;
                return _tab.Val.Value.Convert<TabStopStyle>();
            }
            set
            {
                _tab.Val = value.Convert<W.TabStopValues>();
            }
        }

        public TabStopLeader Leader
        {
            get
            {
                if (_tab.Leader == null) return TabStopLeader.None;
                return _tab.Leader.Value.Convert<TabStopLeader>();
            }
            set
            {
                _tab.Leader = value.Convert<W.TabStopLeaderCharValues>();
            }
        }


        public bool Equals(Tab other)
        {
            return _tab.Equals(other.XETabStop);
        }

        internal W.TabStop XETabStop => _tab;
    }
}
