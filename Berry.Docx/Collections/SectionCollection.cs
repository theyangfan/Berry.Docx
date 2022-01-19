using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx.Collections
{
    public class SectionCollection : IEnumerable
    {
        private IEnumerable<Section> _sections;
        public SectionCollection(IEnumerable<Section> sections)
        {
            _sections = sections;
        }

        public Section this[int index]
        {
            get
            {
                return _sections.ElementAt(index);
            }
        }

        /// <summary>
        /// 返回集合数量
        /// </summary>
        public int Count { get => _sections.Count(); }
        public IEnumerator GetEnumerator()
        {
            return _sections.GetEnumerator();
        }
    }
}
