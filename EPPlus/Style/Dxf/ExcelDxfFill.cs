using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    public class ExcelDxfFill : DxfStyleBase<ExcelDxfFill>
    {
        public ExcelDxfFill(ExcelStyles styles)
            : base(styles)
        {
            PatternColor = new ExcelDxfColor(styles);
            BackgroundColor = new ExcelDxfColor(styles);
        }
        public ExcelFillStyle? PatternType { get; set; }
        /// <summary>
        /// The color of the pattern
        /// </summary>
        public ExcelDxfColor PatternColor { get; internal set; }
        /// <summary>
        /// The background color
        /// </summary>
        public ExcelDxfColor BackgroundColor { get; internal set; }

        protected internal override string Id
        {
            get
            {
                return GetAsString(PatternType) + "|" + GetAsString(PatternColor) + "|" + GetAsString(BackgroundColor);
            }
        }
        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            helper.CreateNode(path);
            SetValueEnum(helper, path + "/d:patternFill/@patternType", PatternType);
            SetValueColor(helper, path + "/d:patternFill/d:fgColor", PatternColor);
            SetValueColor(helper, path + "/d:patternFill/d:bgColor", BackgroundColor);
        }

        protected internal override bool HasValue
        {
            get 
            {
                return PatternType != null ||
                    PatternColor.HasValue ||
                    BackgroundColor.HasValue;
            }
        }
        protected internal override ExcelDxfFill Clone()
        {
            return new ExcelDxfFill(_styles) {PatternType=PatternType, PatternColor=PatternColor.Clone(), BackgroundColor=BackgroundColor.Clone()};
        }
    }
}
