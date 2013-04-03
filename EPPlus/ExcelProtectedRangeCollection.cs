using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml
{
    public class ExcelProtectedRangeCollection : XmlHelper, ICollection<ExcelProtectedRange>
    {
        public ExcelProtectedRangeCollection(XmlNamespaceManager nsm, XmlNode topNode, ExcelWorksheet ws)
            : base(nsm, topNode)
        {
            foreach (XmlNode protectedRangeNode in topNode.SelectNodes("d:protectedRanges/d:protectedRange", nsm))
            {
                if (!(protectedRangeNode is XmlElement))
                    continue;
                _baseCollection.Add(new ExcelProtectedRange(protectedRangeNode.Attributes["name"].Value, new ExcelAddress(protectedRangeNode.Attributes["sqref"].Value)));
            }
        }

        private Collection<ExcelProtectedRange> _baseCollection = new Collection<ExcelProtectedRange>();

        public void Add(ExcelProtectedRange item)
        {
            if (!ExistNode("d:protectedRanges"))
                CreateNode("d:protectedRanges");
            var newNode = CreateNode("d:protectedRanges/d:protectedRange");
            var sqrefAttribute = TopNode.OwnerDocument.CreateAttribute("sqref");
            sqrefAttribute.Value = item.Address.Address;
            newNode.Attributes.Append(sqrefAttribute);
            var nameAttribute = TopNode.OwnerDocument.CreateAttribute("name");
            nameAttribute.Value = item.Name;
            newNode.Attributes.Append(nameAttribute);
            _baseCollection.Add(item);
        }

        public void Clear()
        {
            DeleteNode("d:protectedRanges");
            _baseCollection.Clear();
        }

        public bool Contains(ExcelProtectedRange item)
        {
            return _baseCollection.Contains(item);
        }

        public void CopyTo(ExcelProtectedRange[] array, int arrayIndex)
        {
            _baseCollection.CopyTo(array, arrayIndex);
        }

        public int Count
        {
            get { return _baseCollection.Count; }
        }

        public bool Remove(ExcelProtectedRange item)
        {
            DeleteAllNode("d:protectedRanges/d:protectedRange[@name='" + item.Name + "' and @sqref='" + item.Address.Address + "']");
            if (_baseCollection.Count == 0)
                DeleteNode("d:protectedRanges");
            return _baseCollection.Remove(item);
        }

        public IEnumerator<ExcelProtectedRange> GetEnumerator()
        {
            return _baseCollection.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _baseCollection.GetEnumerator();
        }

        public bool IsReadOnly
        {
            get { return false; }
        }
    }
}
