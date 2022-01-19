using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Berry.Docx.Documents;

namespace Berry.Docx.Collections
{
    /// <summary>
    /// 样式集合
    /// </summary>
    public class StyleCollection : IEnumerable
    {
        private IEnumerable<Style> _styles;
        /// <summary>
        /// 样式集合
        /// </summary>
        /// <param name="styles"></param>
        public StyleCollection(IEnumerable<Style> styles)
        {
            _styles = styles;
        }
        /// <summary>
        /// 返回所以为index的样式
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public Style this[int index]
        {
            get
            {
                return _styles.ElementAt(index);
            }
        }

        /// <summary>
        /// 返回集合数量
        /// </summary>
        public int Count { get => _styles.Count(); }
        /// <summary>
        /// 返回样式枚举器
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            return _styles.GetEnumerator();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public Style FindByName(string name)
        {
            return _styles.Where(s => s.Name.ToLower() == name.ToLower()).FirstOrDefault();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public Style FindByName(string name, StyleType type)
        {
            return _styles.Where(s => s.Name.ToLower() == name.ToLower() && s.Type == type).FirstOrDefault();
        }
    }

    /*
    public class StyleEnumerator : IEnumerator
    {
        private List<Style> _styles;
        int _position = -1;
        public StyleEnumerator(List<Style> styles)
        {
            _styles = styles;
        }

        public object Current
        {
            get
            {
                if (_position == -1)
                    throw new InvalidOperationException();
                if (_position >= _styles.Count)
                    throw new InvalidOperationException();
                return _styles[_position];
            }
        }

        public bool MoveNext()
        {
            if (_position < _styles.Count - 1)
            {
                _position++;
                return true;
            }
            else
            {
                return false;
            }
        }

        public void Reset()
        {
            _position = -1;
        }
    }
    */
}
