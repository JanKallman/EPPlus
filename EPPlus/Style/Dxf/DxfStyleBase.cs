using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    public abstract class DxfStyleBase<T>
    {
        protected ExcelStyles _styles;
        internal DxfStyleBase(ExcelStyles styles)
        {
            _styles = styles;
            AllowChange = false; //Don't touch this value in the styles.xml by default.
        }
        protected internal abstract string Id { get; }
        protected internal abstract bool HasValue{get;}
        protected internal abstract void CreateNodes(XmlHelper helper, string path);
        protected internal abstract T Clone();
        protected void SetValueColor(XmlHelper helper,string path, ExcelDxfColor color)
        {
            if (color != null && color.HasValue)
            {
                if (color.Color != null)
                {
                    SetValue(helper, path + "/@rgb", color.Color.Value.ToArgb().ToString("x"));
                }
                else if (color.Auto != null)
                {
                    SetValueBool(helper, path + "/@auto", color.Auto);
                }
                else if (color.Theme != null)
                {
                    SetValue(helper, path + "/@theme", color.Theme);
                }
                else if (color.Index != null)
                {
                    SetValue(helper, path + "/@index", color.Index);
                }
                if (color.Tint != null)
                {
                    SetValue(helper, path + "/@tint", color.Tint);
                }
            }
        }
        /// <summary>
        /// Same as SetValue but will set first char to lower case.
        /// </summary>
        /// <param name="helper"></param>
        /// <param name="path"></param>
        /// <param name="v"></param>
        protected void SetValueEnum(XmlHelper helper, string path, Enum v)
        {
            if (v == null)
            {
                helper.DeleteNode(path);
            }
            else
            {
                var s = v.ToString();
                s = s.Substring(0, 1).ToLower() + s.Substring(1);
                helper.SetXmlNodeString(path, s);
            }
        }
        protected void SetValue(XmlHelper helper, string path, object v)
        {
            if (v == null)
            {
                helper.DeleteNode(path);
            }
            else
            {
                helper.SetXmlNodeString(path, v.ToString());
            }
        }
        protected void SetValueBool(XmlHelper helper, string path, bool? v)
        {
            if (v == null)
            {
                helper.DeleteNode(path);
            }
            else
            {
                helper.SetXmlNodeBool(path, (bool)v);
            }
        }
        protected internal string GetAsString(object v)
        {
            return (v ?? "").ToString();
        }
        /// <summary>
        /// Is this value allowed to be changed?
        /// </summary>
        protected internal bool AllowChange { get; set; }
    }
}
