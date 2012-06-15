using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    public class ExcelDxfBorderItem : DxfStyleBase<ExcelDxfBorderItem>
    {
        internal ExcelDxfBorderItem(ExcelStyles styles) :
            base(styles)
        {
            Color=new ExcelDxfColor(styles);
        }
        public ExcelBorderStyle? Style { get; set;}
        public ExcelDxfColor Color { get; internal set; }

        protected internal override string Id
        {
            get
            {
                return GetAsString(Style) + "|" + GetAsString(Color);
            }
        }

        protected internal override void CreateNodes(XmlHelper helper, string path)
        {            
            SetValueEnum(helper, path + "/@style", Style);
            SetValueColor(helper, path + "/d:color", Color);
        }
        protected internal override bool HasValue
        {
            get 
            {
                return Style != null || Color.HasValue;
            }
        }
        protected internal override ExcelDxfBorderItem Clone()
        {
            return new ExcelDxfBorderItem(_styles) { Style = Style, Color = Color };
        }
    }
}
