﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Drawing;

namespace OfficeOpenXml.Drawing.Vml
{
    public enum eTextAlignHorizontalVml
    {
        Left,
        Center,
        Right
    }
    public enum eTextAlignVerticalVml
    {
        Top,
        Center,
        Bottom
    }
    /// <summary>
    /// Drawing object used for comments
    /// </summary>
    public class ExcelVmlDrawing : XmlHelper, IRangeID
    {
        public ExcelVmlDrawing(XmlNode topNode, ExcelRangeBase range) :
            base(range.Worksheet.VmlDrawings.NameSpaceManager, topNode)
        {
            Range = range;
            SchemaNodeOrder = new string[] { "fill", "stroke", "shadow", "path", "textbox", "ClientData", "MoveWithCells", "SizeWithCells", "Anchor", "Locked", "AutoFill", "LockText", "TextHAlign", "TextVAlign", "Row", "Column", "Visible" };
        }   
        ExcelRangeBase Range { get; set; }
        public string Id 
        {
            get
            {
                return GetXmlNode("@id");
            }
            set
            {
                SetXmlNode("@id",value);
            }
        }
        const string VERTICAL_ALIGNMENT_PATH="x:ClientData/x:TextVAlign";
        public eTextAlignVerticalVml VerticalAlignment
        {
            get
            {
                switch (GetXmlNode(VERTICAL_ALIGNMENT_PATH))
                {
                    case "Center":
                        return eTextAlignVerticalVml.Center;
                    case "Bottom":
                        return eTextAlignVerticalVml.Bottom;
                    default:
                        return eTextAlignVerticalVml.Top;
                }
            }
            set
            {
                switch (value)
                {
                    case eTextAlignVerticalVml.Center:
                        SetXmlNode(VERTICAL_ALIGNMENT_PATH, "Center");
                        break;
                    case eTextAlignVerticalVml.Bottom:
                        SetXmlNode(VERTICAL_ALIGNMENT_PATH, "Bottom");
                        break;
                    default:
                        DeleteNode(VERTICAL_ALIGNMENT_PATH);
                        break;
                }
            }
        }
        const string HORIZONTAL_ALIGNMENT_PATH="x:ClientData/x:TextHAlign";
        public eTextAlignHorizontalVml HorizontalAlignment
        {
            get
            {
                switch (GetXmlNode(HORIZONTAL_ALIGNMENT_PATH))
                {
                    case "Center":
                        return eTextAlignHorizontalVml.Center;
                    case "Right":
                        return eTextAlignHorizontalVml.Right;
                    default:
                        return eTextAlignHorizontalVml.Left;
                }
            }
            set
            {
                switch (value)
                {
                    case eTextAlignHorizontalVml.Center:
                        SetXmlNode(HORIZONTAL_ALIGNMENT_PATH, "Center");
                        break;
                    case eTextAlignHorizontalVml.Right:
                        SetXmlNode(HORIZONTAL_ALIGNMENT_PATH, "Right");
                        break;
                    default:
                        DeleteNode(HORIZONTAL_ALIGNMENT_PATH);
                        break;
                }
            }
        }
        const string VISIBLE_PATH = "x:ClientData/x:Visible";
        public bool Visible 
        { 
            get
            {
                return (TopNode.SelectSingleNode(VISIBLE_PATH, NameSpaceManager)!=null);
            }
            set
            {
                if (value)
                {
                    CreateNode(VISIBLE_PATH);
                    Style = SetStyle(Style,"visibility", "visible");
                }
                else
                {
                    DeleteNode(VISIBLE_PATH);
                    Style = SetStyle(Style,"visibility", "hidden");
                }                
            }
        }
        const string BACKGROUNDCOLOR_PATH = "@fillcolor";
        const string BACKGROUNDCOLOR2_PATH = "v:fill/@color2";
        public Color BackgroundColor
        {
            get
            {
                string col = GetXmlNode(BACKGROUNDCOLOR_PATH);
                if (col == "")
                {
                    return Color.FromArgb(0xff, 0xff, 0xe1);
                }
                else
                {
                    if(col.StartsWith("#")) col=col.Substring(1,col.Length-1);
                    int res;
                    if (int.TryParse(col,System.Globalization.NumberStyles.AllowHexSpecifier,ExcelWorksheet._ci, out res))
                    {
                        return Color.FromArgb(res);
                    }
                    else
                    {
                        return Color.Empty;
                    }
                }
            }
            set
            {
                string color = "#" + value.ToArgb().ToString("X").Substring(2, 6);
                SetXmlNode(BACKGROUNDCOLOR_PATH, color);
                //SetXmlNode(BACKGROUNDCOLOR2_PATH, color);
            }
        }
        ///// <summary>
        ///// Width of the Comment 
        ///// </summary>
        //public Single Width
        //{
        //    get
        //    {
        //        string v;
        //        GetStyle("width", out v);
        //        if(v.EndsWith("pt"))
        //        {
        //            v = v.Substring(0, v.Length - 2);
        //        }
        //        short ret;
        //        if (short.TryParse(v,System.Globalization.NumberStyles.Any, ExcelWorksheet._ci, out ret))
        //        {
        //            return ret;
        //        }
        //        else
        //        {
        //            return 0;
        //        }
        //    }
        //    set
        //    {
        //        SetStyle("width", value.ToString("N2",ExcelWorksheet._ci) + "pt");
        //    }
        //}
        ///// <summary>
        ///// Height of the Comment 
        ///// </summary>
        //public Single Height
        //{
        //    get
        //    {
        //        string v;
        //        GetStyle("height", out v);
        //        if (v.EndsWith("pt"))
        //        {
        //            v = v.Substring(0, v.Length - 2);
        //        }
        //        short ret;
        //        if (short.TryParse(v, System.Globalization.NumberStyles.Any, ExcelWorksheet._ci, out ret))
        //        {
        //            return ret;
        //        }
        //        else
        //        {
        //            return 0;
        //        }
        //    }
        //    set
        //    {
        //        SetStyle("height", value.ToString("N2", ExcelWorksheet._ci) + "pt");
        //    }
        //}
        const string TEXTBOX_STYLE_PATH = "v:textbox/@style";
        public bool AutoFit
        {
            get
            {
                string value;
                GetStyle(GetXmlNode(TEXTBOX_STYLE_PATH), "mso-fit-shape-to-text", out value);
                return value=="t";
            }
            set
            {                
                SetXmlNode(TEXTBOX_STYLE_PATH, SetStyle(GetXmlNode(TEXTBOX_STYLE_PATH),"mso-fit-shape-to-text", value?"t":"")); 
            }
        }
        ExcelVmlDrawingPosition _from = null;
        public ExcelVmlDrawingPosition From
        {
            get
            {
                if (_from == null)
                {
                    _from = new ExcelVmlDrawingPosition(NameSpaceManager, TopNode.SelectSingleNode("x:ClientData", NameSpaceManager), 0);
                }
                return _from;
            }
        }
        ExcelVmlDrawingPosition _to = null;
        public ExcelVmlDrawingPosition To
        {
            get
            {
                if (_to == null)
                {
                    _to = new ExcelVmlDrawingPosition(NameSpaceManager, TopNode.SelectSingleNode("x:ClientData", NameSpaceManager), 4);
                }
                return _to;
            }
        }
        const string STYLE_PATH = "@style";
        internal string Style
        {
            get
            {
                return GetXmlNode(STYLE_PATH);
            }
            set
            {
                SetXmlNode(STYLE_PATH, value);
            }
        }
        private bool GetStyle(string style, string key, out string value)
        {
            string[]styles = style.Split(';');
            foreach(string s in styles)
            {
                if (s.IndexOf(':') > 0)
                {
                    string[] split = s.Split(':');
                    if (split[0] == key)
                    {
                        value=split[1];
                        return true;
                    }
                }
                else if (s == key)
                {
                    value="";
                    return true;
                }
            }
            value="";
            return false;
        }
        private string SetStyle(string style, string key, string value)
        {
            string[] styles = style.Split(';');
            string newStyle="";
            bool changed = false;
            foreach (string s in styles)
            {
                string[] split = s.Split(':');
                if (split[0].Trim() == key)
                {
                    if (value.Trim() != "") //If blank remove the item
                    {
                        newStyle += key + ':' + value;
                    }
                    changed = true;
                }
                else
                {
                    newStyle += s;
                }
                newStyle += ';';
            }
            if (!changed)
            {
                newStyle += key + ':' + value;
            }
            else
            {
                newStyle = style.Substring(0, style.Length - 1);
            }
            return newStyle;
        }
        #region IRangeID Members

        ulong IRangeID.RangeID
        {
            get
            {
                return ExcelCellBase.GetCellID(Range.Worksheet.SheetID, Range.Start.Row, Range.Start.Column);
            }
            set
            {
                
            }
        }

        #endregion
    }
}
