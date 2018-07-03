﻿/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Drawing;
using System.Globalization;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Collection of Richtext objects
    /// </summary>
    public class ExcelRichTextCollection : XmlHelper, IEnumerable<ExcelRichText>
    {
        List<ExcelRichText> _list = new List<ExcelRichText>();
        ExcelRangeBase _cells=null;
        internal ExcelRichTextCollection(XmlNamespaceManager ns, XmlNode topNode) :
            base(ns, topNode)
        {
            var nl = topNode.SelectNodes("d:r", NameSpaceManager);
            if (nl != null)
            {
                foreach (XmlNode n in nl)
                {
                    _list.Add(new ExcelRichText(ns, n,this));
                }
            }
        }
        internal ExcelRichTextCollection(XmlNamespaceManager ns, XmlNode topNode, ExcelRangeBase cells) :
            this(ns, topNode)
        {
            _cells = cells;
        }        
        /// <summary>
        /// Collection containing the richtext objects
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public ExcelRichText this[int Index]
        {
            get
            {
                var item=_list[Index];
                if(_cells!=null) item.SetCallback(UpdateCells);
                return item;
            }
        }
        /// <summary>
        /// Items in the list
        /// </summary>
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }
        /// <summary>
        /// Add a rich text string
        /// </summary>
        /// <param name="Text">The text to add</param>
        /// <returns></returns>
        public ExcelRichText Add(string Text)
        {
            return Insert(_list.Count, Text);
        }

        /// <summary>
        /// Insert a rich text string at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index at which rich text should be inserted.</param>
        /// <param name="text">The text to insert.</param>
        /// <returns></returns>
        public ExcelRichText Insert(int index, string text)
        {
            ConvertRichtext();
            XmlDocument doc;
            if (TopNode is XmlDocument)
            {
                doc = TopNode as XmlDocument;
            }
            else
            {
                doc = TopNode.OwnerDocument;
            }
            var node = doc.CreateElement("d", "r", ExcelPackage.schemaMain);
            if (index < _list.Count)
            {
                TopNode.InsertBefore(node, TopNode.ChildNodes[index]);
            }
            else
            {
                TopNode.AppendChild(node);
            }
            var rt = new ExcelRichText(NameSpaceManager, node, this);
            if (_list.Count > 0)
            {
                ExcelRichText prevItem = _list[index < _list.Count ? index : _list.Count - 1];
                rt.FontName = prevItem.FontName;
                rt.Size = prevItem.Size;
                if (prevItem.Color.IsEmpty)
                {
                    rt.Color = Color.Black;
                }
                else
                {
                    rt.Color = prevItem.Color;
                }
                rt.PreserveSpace = rt.PreserveSpace;
                rt.Bold = prevItem.Bold;
                rt.Italic = prevItem.Italic;
                rt.UnderLine = prevItem.UnderLine;
            }
            else if (_cells == null)
            {
                rt.FontName = "Calibri";
                rt.Size = 11;
            }
            else
            {
                var style = _cells.Offset(0, 0).Style;
                rt.FontName = style.Font.Name;
                rt.Size = style.Font.Size;
                rt.Bold = style.Font.Bold;
                rt.Italic = style.Font.Italic;
                _cells.IsRichText = true;
            }
            rt.Text = text;
            rt.PreserveSpace = true;
            if (_cells != null)
            {
                rt.SetCallback(UpdateCells);
                UpdateCells();
            }
            _list.Insert(index, rt);
            return rt;
        }

        internal void ConvertRichtext()
        {
            if (_cells == null) return;
            var isRt = _cells.Worksheet._flags.GetFlagValue(_cells._fromRow, _cells._fromCol, CellFlags.RichText);
            if (Count == 1 && isRt == false)
            {
                _cells.Worksheet._flags.SetFlagValue(_cells._fromRow, _cells._fromCol, true, CellFlags.RichText);
                var s = _cells.Worksheet.GetStyleInner(_cells._fromRow, _cells._fromCol);
                //var fnt = cell.Style.Font;
                var fnt = _cells.Worksheet.Workbook.Styles.GetStyleObject(s, _cells.Worksheet.PositionID, ExcelAddressBase.GetAddress(_cells._fromRow, _cells._fromCol)).Font;
                this[0].PreserveSpace = true;
                this[0].Bold = fnt.Bold;
                this[0].FontName = fnt.Name;
                this[0].Italic = fnt.Italic;
                this[0].Size = fnt.Size;
                this[0].UnderLine = fnt.UnderLine;

                int hex;
                if (fnt.Color.Rgb != "" && int.TryParse(fnt.Color.Rgb, NumberStyles.HexNumber, null, out hex))
                {
                    this[0].Color = Color.FromArgb(hex);
                }
            }
        }
        internal void UpdateCells()
        {
            _cells.SetValueRichText(TopNode.InnerXml);
        }
        /// <summary>
        /// Clear the collection
        /// </summary>
        public void Clear()
        {
            _list.Clear();
            TopNode.RemoveAll();
            UpdateCells();
            if (_cells != null) _cells.IsRichText = false;
        }
        /// <summary>
        /// Removes an item at the specific index
        /// </summary>
        /// <param name="Index"></param>
        public void RemoveAt(int Index)
        {
            TopNode.RemoveChild(_list[Index].TopNode);
            _list.RemoveAt(Index);
            if (_cells != null && _list.Count==0) _cells.IsRichText = false;
        }
        /// <summary>
        /// Removes an item
        /// </summary>
        /// <param name="Item"></param>
        public void Remove(ExcelRichText Item)
        {
            TopNode.RemoveChild(Item.TopNode);
            _list.Remove(Item);
            if (_cells != null && _list.Count == 0) _cells.IsRichText = false;
        }
        //public void Insert(int index, string Text)
        //{
        //    _list.Insert(index, item);
        //}
        
        /// <summary>
        /// The text
        /// </summary>
        public string Text
        {
            get
            {
                StringBuilder sb=new StringBuilder();
                foreach (var item in _list)
                {
                    sb.Append(item.Text);
                }
                return sb.ToString();
            }
            set
            {
                if (Count == 0)
                {
                    Add(value);
                }
                else
                {
                    this[0].Text = value;
                    for (int ix = 1; ix < Count; ix++)
                    {
                        RemoveAt(ix);
                    }
                }
            }
        }
        #region IEnumerable<ExcelRichText> Members

        IEnumerator<ExcelRichText> IEnumerable<ExcelRichText>.GetEnumerator()
        {
            return _list.Select(x => { x.SetCallback(UpdateCells); return x; }).GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.Select(x => { x.SetCallback(UpdateCells); return x; }).GetEnumerator();
        }

        #endregion
    }
}
