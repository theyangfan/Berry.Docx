// Copyright (c) theyangfan. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;

using Berry.Docx.Collections;

namespace Berry.Docx.Formatting
{
    /// <summary>
    /// Repersent the list style, Each style has 9 levels.
    /// <para>表示一个多级列表样式，每种样式有9个级别. </para>
    /// </summary>
    public class ListStyle : IEquatable<ListStyle>
    {
        #region Private Members
        private readonly Document _doc;
        private readonly W.AbstractNum _abstractNum;
        #endregion

        #region Constructors
        internal ListStyle(Document doc, W.AbstractNum abstractNum)
        {
            _doc = doc;
            _abstractNum = abstractNum;
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets or sets the style name. The list style name does not exist physically, the name will be invalid when out of document scope.
        /// <para>获取或设置样式名称. 列表样式的名称在物理上不存在，当离开文档作用域后，名称将无效。</para>
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets the list levels collection of the current style.
        /// </summary>
        public ListLevelCollection Levels => new ListLevelCollection(GetLevels());
        #endregion

        #region Public Methods
        /// <summary>
        /// Creates a built-in list style.
        /// <para>创建一个内置的列表样式.</para>
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="style">The built-in list style type.</param>
        /// <returns>The list style.</returns>
        public static ListStyle Create(Document doc, BuiltInListStyle style)
        {
            return new ListStyle(doc, BuiltInListStyleGenerator.Generate(doc, style));
        }

        public bool Equals(ListStyle style)
        {
            return this.AbstractNum.Equals(style.AbstractNum);
        }
        #endregion

        #region Internal Properties
        internal int NumberID => NumberingInstance.NumberID;

        internal int AbstractNumberID => _abstractNum.AbstractNumberId;

        internal W.AbstractNum AbstractNum => _abstractNum;
        internal W.NumberingInstance NumberingInstance
        {
            get
            {
                W.Numbering numbering = _doc.Package.MainDocumentPart.NumberingDefinitionsPart?.Numbering;
                if (numbering != null)
                {
                    W.NumberingInstance num = numbering.Elements<W.NumberingInstance>()
                        .Where(n => n.AbstractNumId.Val == _abstractNum.AbstractNumberId).FirstOrDefault();
                    if (num != null) return num;
                }
                W.NumberingInstance numberingInstance = new W.NumberingInstance()
                {
                    NumberID = IDGenerator.GenerateNumId(_doc)
                };
                numberingInstance.AbstractNumId = new W.AbstractNumId() { Val = _abstractNum.AbstractNumberId };
                return numberingInstance;
            }
        }
        #endregion

        #region Private Methods
        private IEnumerable<ListLevel> GetLevels()
        {
            foreach (W.Level level in _abstractNum.Elements<W.Level>())
            {
                yield return new ListLevel(_doc, _abstractNum, level);
            }
        }
        #endregion
    }
}
