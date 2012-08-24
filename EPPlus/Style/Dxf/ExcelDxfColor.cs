using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    public class ExcelDxfColor : DxfStyleBase<ExcelDxfColor>

    {
        public ExcelDxfColor(ExcelStyles styles) : base(styles)
        {

        }
        public int? Theme { get; set; }
        public int? Index { get; set; }
        public bool? Auto { get; set; }
        public double? Tint { get; set; }
        public Color? Color { get; set; }
        protected internal override string Id
        {
            get { return GetAsString(Theme) + "|" + GetAsString(Index) + "|" + GetAsString(Auto) + "|" + GetAsString(Tint) + "|" + GetAsString(Color==null ? "" : ((Color)Color.Value).ToArgb().ToString("x")); }
        }
        protected internal override ExcelDxfColor Clone()
        {
            return new ExcelDxfColor(_styles) { Theme = Theme, Index = Index, Color = Color, Auto = Auto, Tint = Tint };
        }
        protected internal override bool HasValue
        {
            get
            {
                return Theme != null ||
                       Index != null ||
                       Auto != null ||
                       Tint != null ||
                       Color != null;
            }
        }
        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            throw new NotImplementedException();
        }
    }
}
