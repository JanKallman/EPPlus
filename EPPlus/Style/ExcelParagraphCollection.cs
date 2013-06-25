/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
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
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing;
using System.Drawing;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// A collection of Paragraph objects
    /// </summary>
    public class ExcelParagraphCollection : XmlHelper, IEnumerable<ExcelParagraph>
    {
        List<ExcelParagraph> _list = new List<ExcelParagraph>();
        string _path;
        internal ExcelParagraphCollection(XmlNamespaceManager ns, XmlNode topNode, string path, string[] schemaNodeOrder) :
            base(ns, topNode)
        {
            var nl = topNode.SelectNodes(path + "/a:r", NameSpaceManager);
            SchemaNodeOrder = schemaNodeOrder;
            if (nl != null)
            {
                foreach (XmlNode n in nl)
                {
                    _list.Add(new ExcelParagraph(ns, n, "",schemaNodeOrder));
                }
            }
            _path = path;
        }
        public ExcelParagraph this[int Index]
        {
            get
            {
                return _list[Index];
            }
        }
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
        public ExcelParagraph Add(string Text)
        {
            XmlDocument doc;
            if (TopNode is XmlDocument)
            {
                doc = TopNode as XmlDocument;
            }
            else
            {
                doc = TopNode.OwnerDocument;
            }
            XmlNode parentNode=TopNode.SelectSingleNode(_path, NameSpaceManager);
            if (parentNode == null)
            {
                CreateNode(_path);
            }
            
            var node = doc.CreateElement("a", "r", ExcelPackage.schemaDrawings);
            parentNode.AppendChild(node);
            var childNode = doc.CreateElement("a", "rPr", ExcelPackage.schemaDrawings);
            node.AppendChild(childNode);
            var rt = new ExcelParagraph(NameSpaceManager, node, "", SchemaNodeOrder);
            rt.ComplexFont = "Calibri";
            rt.LatinFont = "Calibri"; 
            rt.Size = 11;

            rt.Text = Text;
            _list.Add(rt);
            return rt;
        }
        public void Clear()
        {
            _list.Clear();
            TopNode.RemoveAll();
        }
        public void RemoveAt(int Index)
        {
            var node = _list[Index].TopNode;
            while (node != null && node.Name != "a:r")
            {
                node = node.ParentNode;
            }
            node.ParentNode.RemoveChild(node);
            _list.RemoveAt(Index);
        }
        public void Remove(ExcelRichText Item)
        {
            TopNode.RemoveChild(Item.TopNode);
        }
        public string Text
        {
            get
            {
                StringBuilder sb = new StringBuilder();
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
                    int count = Count;
                    for (int ix = Count-1; ix > 0; ix--)
                    {
                        RemoveAt(ix);
                    }
                }
            }
        }
        #region IEnumerable<ExcelRichText> Members

        IEnumerator<ExcelParagraph> IEnumerable<ExcelParagraph>.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        #endregion
    }
}
