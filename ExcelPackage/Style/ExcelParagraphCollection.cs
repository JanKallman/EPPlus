using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing;
using System.Drawing;

namespace OfficeOpenXml.Style
{
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
            TopNode.RemoveChild(_list[Index].TopNode);
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
                    for (int ix = 1; ix < Count; ix++)
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
