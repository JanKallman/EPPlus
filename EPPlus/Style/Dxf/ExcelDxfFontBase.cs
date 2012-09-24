using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    public class ExcelDxfFontBase : DxfStyleBase<ExcelDxfFontBase>
    {
        public ExcelDxfFontBase(ExcelStyles styles)
            : base(styles)
        {
            Color = new ExcelDxfColor(styles);
        }
        /// <summary>
        /// Font bold
        /// </summary>
        public bool? Bold
        {
            get;
            set;
        }
        /// <summary>
        /// Font Italic
        /// </summary>
        public bool? Italic
        {
            get;
            set;
        }
        /// <summary>
        /// Font-Strikeout
        /// </summary>
        public bool? Strike { get; set; }
        //public float? Size { get; set; }
        public ExcelDxfColor Color { get; set; }
        //public string Name { get; set; }
        //public int? Family { get; set; }
        ///// <summary>
        ///// Font-Vertical Align
        ///// </summary>
        //public ExcelVerticalAlignmentFont? VerticalAlign
        //{
        //    get;
        //    set;
        //}

        public ExcelUnderLineType? Underline { get; set; }

        protected internal override string Id
        {
            get
            {
                return GetAsString(Bold) + "|" + GetAsString(Italic) + "|" + GetAsString(Strike) + "|" + (Color ==null ? "" : Color.Id) + "|" /*+ GetAsString(VerticalAlign) + "|"*/ + GetAsString(Underline);
            }
        }

        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            helper.CreateNode(path);
            SetValueBool(helper, path + "/d:b/@val", Bold);
            SetValueBool(helper, path + "/d:i/@val", Italic);
            SetValueBool(helper, path + "/d:strike", Strike);
            SetValue(helper, path + "/d:u/@val", Underline);
            SetValueColor(helper, path + "/d:color", Color);
        }
        protected internal override bool HasValue
        {
            get
            {
                return Bold != null ||
                       Italic != null ||
                       Strike != null ||
                       Underline != null ||
                       Color.HasValue;
            }
        }
        protected internal override ExcelDxfFontBase Clone()
        {
            return new ExcelDxfFontBase(_styles) { Bold = Bold, Color = Color.Clone(), Italic = Italic, Strike = Strike, Underline = Underline };
        }
    }
}
