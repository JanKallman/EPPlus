/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;
using System.Globalization;
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Position of the axis.
    /// </summary>
    public enum eAxisPosition
    {
        Left = 0,
        Bottom = 1,
        Right = 2,
        Top = 3
    }
    /// <summary>
    /// Position of the Y-Axis
    /// </summary>
    public enum eYAxisPosition
    {
        Left = 0,
        Right = 2,
    }
    /// <summary>
    /// Position of the X-Axis
    /// </summary>
    public enum eXAxisPosition
    {
        Bottom = 1,
        Top = 3
    }
    /// <summary>
    /// Axis orientaion
    /// </summary>
    public enum eAxisOrientation
    {
        MaxMin,
        MinMax
    }
    /// <summary>
    /// How the axis are crossed
    /// </summary>
    public enum eCrossBetween
    {
        /// <summary>
        /// Specifies the value axis shall cross the category axis between data markers
        /// </summary>
        Between,
        /// <summary>
        /// Specifies the value axis shall cross the category axis at the midpoint of a category.
        /// </summary>
        MidCat
    }
    /// <summary>
    /// Where the axis cross. 
    /// </summary>
    public enum eCrosses
    {
        /// <summary>
        /// (Axis Crosses at Zero) The category axis crosses at the zero point of the valueaxis (if possible), or the minimum value (if theminimum is greater than zero) or the maximum (if the maximum is less than zero).
        /// </summary>
        AutoZero,
        /// <summary>
        /// The axis crosses at the maximum value
        /// </summary>
        Max,
        /// <summary>
        /// (Axis crosses at the minimum value of the chart.
        /// </summary>
        Min
    }
    /// <summary>
    /// Tickmarks
    /// </summary>
    public enum eAxisTickMark
    {
        /// <summary>
        /// Specifies the tick marks shall cross the axis. 
        /// </summary>
        Cross,   
        /// <summary>
        /// Specifies the tick marks shall be inside the plot area. 
        /// </summary>
        In,     
        /// <summary>
        /// Specifies there shall be no tick marks.
        /// </summary>
        None,    
        /// <summary>
        /// Specifies the tick marks shall be outside the plot area.
        /// </summary>
        Out
    }
    public enum eBuildInUnits : long
    {
        hundreds = 100,
        thousands = 1000,
        tenThousands = 10000,
        hundredThousands = 100000,
        millions = 1000000,
        tenMillions = 10000000,
        hundredMillions = 100000000,
        billions = 1000000000,
        trillions = 1000000000000
    }
    /// <summary>
    /// An axis for a chart
    /// </summary>
    public sealed class ExcelChartAxis : XmlHelper
    {
        /// <summary>
        /// Type of axis
        /// </summary>
        internal enum eAxisType
        {
            /// <summary>
            /// Value axis
            /// </summary>
            Val,
            /// <summary>
            /// Category axis
            /// </summary>
            Cat,
            /// <summary>
            /// Date axis
            /// </summary>
            Date,
            /// <summary>
            /// Series axis
            /// </summary>
            Serie
        }
        internal ExcelChartAxis(XmlNamespaceManager nameSpaceManager, XmlNode topNode) :
            base(nameSpaceManager, topNode)
        {
            SchemaNodeOrder = new string[] { "axId", "scaling", "logBase", "orientation", "max", "min", "delete", "axPos", "majorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crossesAt", "crosses", "crossBetween", "auto", "lblOffset", "majorUnit", "minorUnit", "dispUnits", "spPr", "txPr" };
        }
        internal string Id
        {
            get
            {
                return GetXmlNodeString("c:axId/@val");
            }
        }
        const string _majorTickMark = "c:majorTickMark/@val";
        /// <summary>
        /// majorTickMark 
        /// This element specifies the major tick marks for the axis. 
        /// </summary>
        public eAxisTickMark MajorTickMark
        {
            get
            {
                var v=GetXmlNodeString(_majorTickMark);
                if(string.IsNullOrEmpty(v))
                {
                    return eAxisTickMark.Cross;
                }
                else
                {
                    try
                    {
                        return (eAxisTickMark)Enum.Parse( typeof( eAxisTickMark ), v );
                    }
                    catch
                    {
                        return eAxisTickMark.Cross;
                    }
                }
            }
            set
            {
                SetXmlNodeString( _majorTickMark, value.ToString().ToLower(CultureInfo.InvariantCulture) );
            }
        }

        const string _minorTickMark = "c:minorTickMark/@val";
        /// <summary>
        /// minorTickMark 
        /// This element specifies the minor tick marks for the axis. 
        /// </summary>
        public eAxisTickMark MinorTickMark
        {
            get
            {
                var v=GetXmlNodeString(_minorTickMark);
                if(string.IsNullOrEmpty(v))
                {
                    return eAxisTickMark.Cross;
                }
                else
                {
                    try
                    {
                        return (eAxisTickMark)Enum.Parse(typeof(eAxisTickMark), v );
                    }
                    catch
                    {
                        return eAxisTickMark.Cross;
                    }
                }
            }
            set
            {
                SetXmlNodeString(_minorTickMark, value.ToString().ToLower(CultureInfo.InvariantCulture));
            }
        }
         /// <summary>
        /// Type of axis
        /// </summary>
        internal eAxisType AxisType
        {
            get
            {
                try
                {
                    return (eAxisType)Enum.Parse(typeof(eAxisType), TopNode.LocalName.Substring(0,3), true);
                }
                catch
                {
                    return eAxisType.Val;
                }
            }
        }
        private string AXIS_POSITION_PATH = "c:axPos/@val";
        /// <summary>
        /// Where the axis is located
        /// </summary>
        public eAxisPosition AxisPosition
        {
            get
            {                
                switch(GetXmlNodeString(AXIS_POSITION_PATH))
                {
                    case "b":
                        return eAxisPosition.Bottom;
                    case "r":
                        return eAxisPosition.Right;
                    case "t":
                        return eAxisPosition.Top;
                    default: 
                        return eAxisPosition.Left;
                }
            }
            internal set
            {
                SetXmlNodeString(AXIS_POSITION_PATH, value.ToString().ToLower(CultureInfo.InvariantCulture).Substring(0,1));
            }
        }
        const string _crossesPath = "c:crosses/@val";
        /// <summary>
        /// Where the axis cross
        /// </summary>
        public eCrosses Crosses
        {
            get
            {
                var v=GetXmlNodeString(_crossesPath);
                if (string.IsNullOrEmpty(v))
                {
                    return eCrosses.AutoZero;
                }
                else
                {
                    try
                    {
                        return (eCrosses)Enum.Parse(typeof(eCrosses), v, true);
                    }
                    catch
                    {
                        return eCrosses.AutoZero;
                    }
                }
            }
            set
            {
                var v = value.ToString();
                v = v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1, v.Length - 1);
                SetXmlNodeString(_crossesPath, v);
            }

        }
        const string _crossBetweenPath = "c:crossBetween/@val";
        /// <summary>
        /// How the axis are crossed
        /// </summary>
        public eCrossBetween CrossBetween
        {
            get
            {
                var v=GetXmlNodeString(_crossBetweenPath);
                if(string.IsNullOrEmpty(v))
                {
                    return eCrossBetween.Between;
                }
                else
                {
                    try
                    {
                        return (eCrossBetween)Enum.Parse(typeof(eCrossBetween), v, true);
                    }
                    catch
                    {
                        return eCrossBetween.Between;
                    }
                }
            }
            set
            {
                var v = value.ToString();
                v = v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1);
                SetXmlNodeString(_crossBetweenPath, v);
            }
        }
        const string _crossesAtPath = "c:crossesAt/@val";
        /// <summary>
        /// The value where the axis cross. 
        /// Null is automatic
        /// </summary>
        public double? CrossesAt
        {
            get
            {
                return GetXmlNodeDoubleNull(_crossesAtPath);
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_crossesAtPath);
                }
                else
                {
                    SetXmlNodeString(_crossesAtPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        const string _formatPath = "c:numFmt/@formatCode";
        /// <summary>
        /// Numberformat
        /// </summary>
        public string Format 
        {
            get
            {
                return GetXmlNodeString(_formatPath);
            }
            set
            {
                SetXmlNodeString(_formatPath,value);
                if(string.IsNullOrEmpty(value))
                {
                    SourceLinked = true;
                }
                else
                {
                    SourceLinked = false;
                }
            }
        }
        const string _sourceLinkedPath = "c:numFmt/@sourceLinked";
        public bool SourceLinked
        {
            get
            {
                return GetXmlNodeBool(_sourceLinkedPath);
            }
            set
            {
                SetXmlNodeBool(_sourceLinkedPath, value);
            }
        }
        const string _lblPos = "c:tickLblPos/@val";
        /// <summary>
        /// Position of the labels
        /// </summary>
        public eTickLabelPosition LabelPosition
        {
            get
            {
                var v=GetXmlNodeString(_lblPos);
                if (string.IsNullOrEmpty(v))
                {
                    return eTickLabelPosition.NextTo;
                }
                else
                {
                    try
                    {
                        return (eTickLabelPosition)Enum.Parse(typeof(eTickLabelPosition), v, true);
                    }
                    catch
                    {
                        return eTickLabelPosition.NextTo;
                    }
                }
            }
            set
            {
                string lp = value.ToString();
                SetXmlNodeString(_lblPos, lp.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + lp.Substring(1, lp.Length - 1));
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Access to fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(NameSpaceManager, TopNode, "c:spPr");
                }
                return _fill;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Access to border properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(NameSpaceManager, TopNode, "c:spPr/a:ln");
                }
                return _border;
            }
        }
        ExcelTextFont _font = null;
        /// <summary>
        /// Access to font properties
        /// </summary>
        public ExcelTextFont Font
        {
            get
            {
                if (_font == null)
                {
                    if (TopNode.SelectSingleNode("c:txPr", NameSpaceManager) == null)
                    {
                        CreateNode("c:txPr/a:bodyPr");
                        CreateNode("c:txPr/a:lstStyle");
                    }
                    _font = new ExcelTextFont(NameSpaceManager, TopNode, "c:txPr/a:p/a:pPr/a:defRPr", new string[] { "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" });
                }
                return _font;
            }
        }
        /// <summary>
        /// If the axis is deleted
        /// </summary>
        public bool Deleted 
        {
            get
            {
                return GetXmlNodeBool("c:delete/@val");
            }
            set
            {
                SetXmlNodeBool("c:delete/@val", value);
            }
        }
        const string _ticLblPos_Path = "c:tickLblPos/@val";
        /// <summary>
        /// Position of the Lables
        /// </summary>
        public eTickLabelPosition TickLabelPosition 
        {
            get
            {
                string v = GetXmlNodeString(_ticLblPos_Path);
                if (v == "")
                {
                    return eTickLabelPosition.None;
                }
                else
                {
                    return (eTickLabelPosition)Enum.Parse(typeof(eTickLabelPosition), v, true);
                }
            }
            set
            {
                string v = value.ToString();
                v=v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1, v.Length - 1);
                SetXmlNodeString(_ticLblPos_Path,v);
            }
        }
        const string _displayUnitPath = "c:dispUnits/c:builtInUnit/@val";
        const string _custUnitPath = "c:dispUnits/c:custUnit/@val";
        public double DisplayUnit
        {
            get
            {
                string v = GetXmlNodeString(_displayUnitPath);
                if (string.IsNullOrEmpty(v))
                {
                    var c = GetXmlNodeDoubleNull(_custUnitPath);
                    if (c == null)
                    {
                        return 0;
                    }
                    else
                    {
                        return c.Value;
                    }
                }
                else
                {
                    try
                    {
                        return (double)(long)Enum.Parse(typeof(eBuildInUnits), v, true);
                    }
                    catch
                    {
                        return 0;
                    }
                }
            }
            set
            {
                if (AxisType == eAxisType.Val && value>=0)
                {
                    foreach(var v in Enum.GetValues(typeof(eBuildInUnits)))
                    {
                        if((double)(long)v==value)
                        {
                            DeleteNode(_custUnitPath);
                            SetXmlNodeString(_displayUnitPath, ((eBuildInUnits)value).ToString());
                            return;
                        }
                    }
                    DeleteNode(_displayUnitPath);
                    if(value!=0)
                    {
                        SetXmlNodeString(_custUnitPath, value.ToString(CultureInfo.InvariantCulture));
                    }
                }
            }
        }        
        ExcelChartTitle _title = null;
        /// <summary>
        /// Chart axis title
        /// </summary>
        public ExcelChartTitle Title
        {
            get
            {
                if (_title == null)
                {
                    var node = TopNode.SelectSingleNode("c:title", NameSpaceManager);
                    if (node == null)
                    {
                        CreateNode("c:title");
                        node = TopNode.SelectSingleNode("c:title", NameSpaceManager);
                        node.InnerXml = "<c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:r><a:t /></a:r></a:p></c:rich></c:tx><c:layout /><c:overlay val=\"0\" />";
                    }
                    _title = new ExcelChartTitle(NameSpaceManager, TopNode);
                }
                return _title;
            }            
        }
        #region "Scaling"
        const string _minValuePath = "c:scaling/c:min/@val";
        /// <summary>
        /// Minimum value for the axis.
        /// Null is automatic
        /// </summary>
        public double? MinValue
        {
            get
            {
                return GetXmlNodeDoubleNull(_minValuePath);
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_minValuePath);
                }
                else
                {
                    SetXmlNodeString(_minValuePath, ((double)value).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        const string _maxValuePath = "c:scaling/c:max/@val";
        /// <summary>
        /// Max value for the axis.
        /// Null is automatic
        /// </summary>
        public double? MaxValue
        {
            get
            {
                return GetXmlNodeDoubleNull(_maxValuePath);
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_maxValuePath);
                }
                else
                {
                    SetXmlNodeString(_maxValuePath, ((double)value).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        const string _majorUnitPath = "c:majorUnit/@val";
        const string _majorUnitCatPath = "c:tickLblSkip/@val";
        /// <summary>
        /// Major unit for the axis.
        /// Null is automatic
        /// </summary>
        public double? MajorUnit
        {
            get
            {
                if (AxisType == eAxisType.Cat)
                {
                    return GetXmlNodeDoubleNull(_majorUnitCatPath);
                }
                else
                {
                    return GetXmlNodeDoubleNull(_majorUnitPath);
                }
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_majorUnitPath);
                    DeleteNode(_majorUnitCatPath);
                }
                else
                {
                    if (AxisType == eAxisType.Cat)
                    {
                        SetXmlNodeString(_majorUnitCatPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        SetXmlNodeString(_majorUnitPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                    }
                }
            }
        }
        const string _minorUnitPath = "c:minorUnit/@val";
        const string _minorUnitCatPath = "c:tickMarkSkip/@val";
        /// <summary>
        /// Minor unit for the axis.
        /// Null is automatic
        /// </summary>
        public double? MinorUnit
        {
            get
            {
                if (AxisType == eAxisType.Cat)
                {
                    return GetXmlNodeDoubleNull(_minorUnitCatPath);
                }
                else
                {
                    return GetXmlNodeDoubleNull(_minorUnitPath);
                }
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_minorUnitPath);
                    DeleteNode(_minorUnitCatPath);
                }
                else
                {
                    if (AxisType == eAxisType.Cat)
                    {
                        SetXmlNodeString(_minorUnitCatPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        SetXmlNodeString(_minorUnitPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                    }
                }
            }
        }
        const string _logbasePath = "c:scaling/c:logBase/@val";
        /// <summary>
        /// The base for a logaritmic scale
        /// Null for a normal scale
        /// </summary>
        public double? LogBase
        {
            get
            {
                return GetXmlNodeDoubleNull(_logbasePath);
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_logbasePath);
                }
                else
                {
                    double v = ((double)value);
                    if (v < 2 || v > 1000)
                    {
                        throw(new ArgumentOutOfRangeException("Value must be between 2 and 1000"));
                    }
                    SetXmlNodeString(_logbasePath, v.ToString("0.0", CultureInfo.InvariantCulture));
                }
            }
        }
        const string _orientationPath = "c:scaling/c:orientation/@val";
        /// <summary>
        /// 
        /// </summary>
        public eAxisOrientation Orientation
        {
            get
            {
                string v = GetXmlNodeString(_orientationPath);
                if (v == "")
                {
                    return eAxisOrientation.MinMax;
                }
                else
                {
                    return (eAxisOrientation)Enum.Parse(typeof(eAxisOrientation), v, true);
                }
            }
            set
            {
                string s=value.ToString();
                s=s.Substring(0,1).ToLower(CultureInfo.InvariantCulture) + s.Substring(1,s.Length-1);
                SetXmlNodeString(_orientationPath, s);
            }
        }
        #endregion
    }
}
