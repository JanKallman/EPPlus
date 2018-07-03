﻿using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml
{
    public class ExcelProtectedRangeCollection : XmlHelper, IEnumerable<ExcelProtectedRange>
    {
        internal ExcelProtectedRangeCollection(XmlNamespaceManager nsm, XmlNode topNode, ExcelWorksheet ws)
            : base(nsm, topNode)
        {
            SchemaNodeOrder = ws.SchemaNodeOrder; //Fixed issue 15385
            foreach (XmlNode protectedRangeNode in topNode.SelectNodes("d:protectedRanges/d:protectedRange", nsm))
            {
                if (!(protectedRangeNode is XmlElement))
                    continue;
                _baseList.Add(new ExcelProtectedRange(protectedRangeNode.Attributes["name"].Value, new ExcelAddress(SqRefUtility.FromSqRefAddress(protectedRangeNode.Attributes["sqref"].Value)), nsm, topNode));
            }
        }

        private List<ExcelProtectedRange> _baseList = new List<ExcelProtectedRange>();

        public ExcelProtectedRange Add(string name, ExcelAddress address)
        {
            if (!ExistNode("d:protectedRanges"))
            {
                CreateNode("d:protectedRanges");
            }
            foreach(var pr in _baseList)
            {
                if(name.Equals(pr.Name,StringComparison.CurrentCultureIgnoreCase))
                {
                    throw (new InvalidOperationException($"A protected range with the namn {name} already exists"));
                }
            }
            var newNode = TopNode.OwnerDocument.CreateElement("protectedRange", ExcelPackage.schemaMain);
            TopNode.SelectSingleNode("d:protectedRanges",NameSpaceManager).AppendChild(newNode);
            var item = new ExcelProtectedRange(name, address, base.NameSpaceManager, newNode);
            _baseList.Add(item);
            return item;
        }

        public void Clear()
        {
            DeleteNode("d:protectedRanges");
            _baseList.Clear();
        }

        public bool Contains(ExcelProtectedRange item)
        {
            return _baseList.Contains(item);
        }

        public void CopyTo(ExcelProtectedRange[] array, int arrayIndex)
        {
            _baseList.CopyTo(array, arrayIndex);
        }

        public int Count
        {
            get { return _baseList.Count; }
        }

        public bool Remove(ExcelProtectedRange item)
        {
            DeleteAllNode("d:protectedRanges/d:protectedRange[@name='" + item.Name + "' and @sqref='" + item.Address.Address + "']");
            if (_baseList.Count == 0)
                DeleteNode("d:protectedRanges");
            return _baseList.Remove(item);
        }

        public int IndexOf(ExcelProtectedRange item)
        {
            return _baseList.IndexOf(item);
        }

        public void RemoveAt(int index)
        {
            _baseList.RemoveAt(index);
        }

        public ExcelProtectedRange this[int index]
        {
            get
            {
                return _baseList[index];
            }
        }

        IEnumerator<ExcelProtectedRange> IEnumerable<ExcelProtectedRange>.GetEnumerator()
        {
            return _baseList.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _baseList.GetEnumerator();
        }
    }
}
