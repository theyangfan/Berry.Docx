using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Represent the custom tab stops collection that supports enumeration.
    /// <para>表示一个支持枚举的制表位集合。</para>
    /// </summary>
    public class TabStops : IEnumerable<TabStop>
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.Paragraph _ownerParagraph;
        private readonly W.Style _ownerStyle;
        private readonly IEnumerable<TabStop> _tabStops;
        #endregion

        #region Constructors
        internal TabStops(Document doc, W.Paragraph paragraph)
        {
            _doc = doc;
            _ownerParagraph = paragraph;
            _tabStops = GetTabs();
        }

        internal TabStops(Document doc, W.Style style)
        {
            _doc = doc;
            _ownerStyle = style;
            _tabStops = GetTabs();
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets the tab stop at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        /// <returns>The tab stop.</returns>
        public TabStop this[int index] => _tabStops.ElementAt(index);

        /// <summary>
        /// Gets the tab stops count.
        /// </summary>
        public int Count => _tabStops.Count();
        #endregion

        #region Public Methods
        /// <summary>
        /// Adds a tab stop.
        /// </summary>
        /// <param name="pos">The position.</param>
        /// <param name="style">The tab stop style.</param>
        /// <param name="leader">The tab stop leader.</param>
        public void Add(float pos, TabStopStyle style, TabStopLeader leader)
        {
            TabStop tab = new TabStop(pos, style, leader);
            Add(tab);
        }

        /// <summary>
        /// Adds a tab stop.
        /// </summary>
        /// <param name="tab">The tab stop.</param>
        public void Add(TabStop tab)
        {
            if (_tabStops.Contains(tab)) return;
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

        /// <summary>
        /// Clears all tab stops.
        /// </summary>
        public void Clear()
        {
            if (_ownerParagraph?.ParagraphProperties?.Tabs != null)
                _ownerParagraph.ParagraphProperties.Tabs = null;
            if (_ownerStyle != null)
            {
                if(_ownerStyle.StyleParagraphProperties?.Tabs != null)
                    _ownerStyle.StyleParagraphProperties.Tabs = null;
                // clear tabs in base style.
                W.Style baseStyle = _ownerStyle.GetBaseStyle(_doc);
                if (baseStyle != null)
                {
                    new TabStops(_doc, baseStyle).Clear();
                }
            }
        }

        /// <summary>
        /// Returns an enumerator that iterates through the tab stops collection.
        /// </summary>
        /// <returns></returns>
        public IEnumerator<TabStop> GetEnumerator()
        {
            return _tabStops.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        #endregion

        #region Private Methods
        private IEnumerable<TabStop> GetTabs()
        {
            if (_ownerParagraph != null)
            {
                if (_ownerParagraph.ParagraphProperties?.Tabs != null)
                {
                    foreach (W.TabStop tab in _ownerParagraph.ParagraphProperties.Tabs.Elements<W.TabStop>())
                        yield return new TabStop(tab);
                }
                foreach (W.TabStop tab in GetStyleTabsRecursively(_doc, _ownerParagraph.GetStyle(_doc)))
                    yield return new TabStop(tab);
            }
            else if (_ownerStyle != null)
            {
                foreach (W.TabStop tab in GetStyleTabsRecursively(_doc, _ownerStyle))
                    yield return new TabStop(tab);
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
        #endregion
    }

    /// <summary>
    /// Represent a single custom tab stop.
    /// <para>表示单个制表位.</para>
    /// </summary>
    public class TabStop : IEquatable<TabStop>
    {
        #region Private Members
        private readonly W.TabStop _tab;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of TabStop Class.
        /// </summary>
        public TabStop() : this(new W.TabStop() { Position = 0, Val = W.TabStopValues.Clear, Leader = W.TabStopLeaderCharValues.None })
        {
        }

        /// <summary>
        /// Initializes a new instance of TabStop Class with specified properties.
        /// </summary>
        /// <param name="pos">The position.</param>
        /// <param name="style">The tab stop style.</param>
        /// <param name="leader">The tab stop leader.</param>
        public TabStop(float pos, TabStopStyle style, TabStopLeader leader) : this()
        {
            Position = pos;
            Style = style;
            Leader = leader;
        }

        internal TabStop(W.TabStop tab)
        {
            _tab = tab;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets or sets the tab stop position (in points).
        /// </summary>
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

        /// <summary>
        /// Gets or sets the tab stop style.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the tab stop leader.
        /// </summary>
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
        #endregion

        #region Public Methods
        /// <summary>
        /// Determines whether the specified TabStop is equal to the current TabStop.
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(TabStop other)
        {
            return _tab.Equals(other.XETabStop);
        }
        #endregion

        #region Internal Properties
        internal W.TabStop XETabStop => _tab;
        #endregion

    }
}
