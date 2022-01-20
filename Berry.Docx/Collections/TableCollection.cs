using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using P = DocumentFormat.OpenXml.Packaging;

using Berry.Docx.Documents;

namespace Berry.Docx.Collections
{
    public class TableCollection : IEnumerable
    {
        private O.OpenXmlElement _owner;
        private IEnumerable<Table> _tables;
        internal TableCollection(O.OpenXmlElement owner, IEnumerable<Table> tables)
        {
            _owner = owner;
            _tables = tables;
        }

        public Table this[int index] => _tables.ElementAt(index);

        /// <summary>
        /// 返回集合数量
        /// </summary>
        public int Count => _tables.Count();

        public bool Contains(Table table)
        {
            return _tables.Contains(table);
        }

        /// <summary>
        /// 在集合末尾添加段落
        /// </summary>
        /// <param name="table">段落</param>
        public void Add(Table table)
        {
            W.Table newTable = table.XElement as W.Table;
            if (_tables.Count() == 0)
            {
                if(_owner is W.Body)
                {
                    _owner.InsertBefore(newTable, _owner.LastChild);
                    return;
                }
                _owner.AppendChild(newTable);
            }
            else
            {
                _tables.Last().XElement.InsertAfterSelf(newTable);
            }
        }

        /// <summary>
        /// 返回段落在集合中从零开始的索引
        /// </summary>
        /// <param name="table">段落</param>
        /// <returns></returns>
        public int IndexOf(Table table)
        {
            return _tables.ToList().IndexOf(table);
        }

        /// <summary>
        /// 在集合指定位置插入段落
        /// </summary>
        /// <param name="table">段落</param>
        /// <param name="index">段落位置，从零开始的索引</param>
        public void InsertAt(Table table, int index)
        {
            W.Table newTable = table.XElement as W.Table;
            if (_tables.Count() == 0)
            {
                if (index == 0)
                {
                    if (_owner is W.Body)
                    {
                        _owner.InsertBefore(newTable, _owner.LastChild);
                        return;
                    }
                    _owner.AppendChild(newTable);
                }
                else
                {
                    throw new ArgumentOutOfRangeException("index", index, "索引超出范围, 必须为非负值并小于集合大小。");
                }
            }
            else
            {
                _tables.ElementAt(index).XElement.InsertBeforeSelf(newTable);
            }
                
        }

        /// <summary>
        /// 移除段落
        /// </summary>
        /// <param name="table">表格</param>
        public void Remove(Table table)
        {
            if (!Contains(table)) return;
            table.Remove();
        }

        /// <summary>
        /// 移除指定位置处的段落
        /// </summary>
        /// <param name="index">段落位置，从零开始的索引</param>
        public void RemoveAt(int index)
        {
            _tables.ElementAt(index).Remove();
        }

        /// <summary>
        /// 移除所有段落
        /// </summary>
        public void Clear()
        {
            foreach (Table table in _tables)
                table.Remove();
        }

        public IEnumerator GetEnumerator()
        {
            return _tables.GetEnumerator();
        }

        
    }
}
