using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table
{
    public class ExcelPivotTableDataFieldSettings  : XmlHelper
    {
        ExcelPivotTableField _field;
        public ExcelPivotTableDataFieldSettings(XmlNamespaceManager ns, XmlNode topNode,ExcelPivotTableField field, int index) :
            base(ns, topNode)
        {
            Index = index;
            BaseField = 0;
            BaseItem = 0;
            _field = field;
        }
        public int Index 
        { 
            get
            {
                return GetXmlNodeInt("@fld");
            }
            set
            {
                SetXmlNodeString("@fld",value.ToString());
            }
        }
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
            set
            {
                SetXmlNodeString("@name", value);
            }
        }
        public int BaseField
        {
            get
            {
                return GetXmlNodeInt("@baseField");
            }
            set
            {
                SetXmlNodeString("@baseField", value.ToString());
            }
        }
        public int BaseItem
        {
            get
            {
                return GetXmlNodeInt("@baseItem");
            }
            set
            {
                SetXmlNodeString("@baseItem", value.ToString());
            }
        }
        public int NumFmtId
        {
            get
            {
                return GetXmlNodeInt("@numFmtId");
            }
            set
            {
                SetXmlNodeString("@numFmtId", value.ToString());
            }
        }        
        public DataFieldFunctions Function
        {
            get
            {
                string s=GetXmlNodeString("@subtotal");
                if(s=="")
                {
                    return DataFieldFunctions.None;
                }
                else
                {
                    return (DataFieldFunctions)Enum.Parse(typeof(DataFieldFunctions), s, true);
                }
            }
            set
            {
                string v;
                switch(value)
                {
                    case DataFieldFunctions.None:
                        DeleteNode("@subtotal");
                        return;
                    case DataFieldFunctions.CountNums:
                        v="CountNums";
                        break;
                    case DataFieldFunctions.StdDev:
                        v="stdDev";
                        break;
                    case DataFieldFunctions.StdDevP:
                        v="stdDevP";
                        break;
                    default:
                        v=value.ToString().ToLower();
                        break;
                }                
                SetXmlNodeString("@subtotal", v);
            }
        }
    }
}
