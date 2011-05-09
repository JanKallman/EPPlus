using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotTableFieldItem : XmlHelper
    {
        ExcelPivotTableField _field;
        public ExcelPivotTableFieldItem(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTableField field) :
            base(ns, topNode)
        {
           _field = field;
        }
        public string Text
        {
            get
            {
                return GetXmlNodeString("@n");
            }
            set
            {
                if(string.IsNullOrEmpty(value))
                {
                    DeleteNode("@n");
                    return;
                }
                foreach (var item in _field.Items)
                {
                    if (item.Text == value)
                    {
                        throw(new ArgumentException("Duplicate Text"));
                    }
                }
                SetXmlNodeString("@n", value);
            }
        }
        internal int X
        {
            get
            {
                return GetXmlNodeInt("@x"); 
            }
        }
        internal string T
        {
            get
            {
                return GetXmlNodeString("@t");
            }
        }
    }
}
