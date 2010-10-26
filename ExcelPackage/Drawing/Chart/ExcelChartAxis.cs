/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 *
 * See http://www.codeplex.com/EPPlus for details.
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * The GNU General Public License can be viewed at http://www.opensource.org/licenses/gpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 * 
 * The code for this project may be used and redistributed by any means PROVIDING it is 
 * not sold for profit without the author's written consent, and providing that this notice 
 * and the author's name and all copyright notices remain intact.
 * 
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * 
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 * ******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;
using System.Globalization;
namespace OfficeOpenXml.Drawing.Chart
{
    public enum eAxisPosition
    {
        Left = 0,
        Bottom = 1,
        Right = 2,
        Top = 3
    }
    public enum eYAxisPosition
    {
        Left = 0,
        Right = 2,
    }
    public enum eXAxisPosition
    {
        Bottom = 1,
        Top = 3
    }
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
    /// An axis for a chart
    /// </summary>
    public sealed class ExcelChartAxis : XmlHelper
    {
        internal enum eAxisType
        {
            Val,
            Cat,
            Date
        }
        internal ExcelChartAxis(XmlNamespaceManager nameSpaceManager, XmlNode topNode) :
            base(nameSpaceManager, topNode)
        {
            SchemaNodeOrder = new string[] { "axId", "scaling", "logBase", "orientation", "max", "min", "delete", "axPos", "majorGridlines", "numFmt", "tickLblPos","spPr","txPr", "crossAx", "crossesAt", "crosses", "crossBetween","auto", "lblOffset","majorUnit","minorUnit", "spPr", "txPr" };
        }
        internal string Id
        {
            get
            {
                return GetXmlNodeString("c:axId/@val");
            }
        }
        internal eAxisType AxisType
        {
            get
            {
                try
                {
                    return (eAxisType)Enum.Parse(typeof(eAxisType), TopNode.LocalName, true);
                }
                catch
                {
                    return eAxisType.Val;
                }
            }
        }
        private string AXIS_POSITION_PATH = "c:axPos/@val";
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
                SetXmlNodeString(AXIS_POSITION_PATH, value.ToString().ToLower().Substring(0,1));
            }
        }
        const string _crossesPath = "c:crosses/@val";
        public eCrosses Crosses
        {
            get
            {
                return (eCrosses)Enum.Parse(typeof(eCrosses), GetXmlNodeString(_crossesPath), true);
            }
            set
            {
                var v = value.ToString();
                v = v.Substring(1).ToLower() + v.Substring(1, v.Length - 1);
                SetXmlNodeString(_crossesPath, v);
            }

        }
        const string _crossBetweenPath = "c:crossBetween/@val";        
        public eCrossBetween CrossBetween
        {
            get
            {
                return (eCrossBetween)Enum.Parse(typeof(eCrossBetween), GetXmlNodeString(_crossBetweenPath), true);
            }
            set
            {
                var v = value.ToString();
                v = v.Substring(1).ToLower() + v.Substring(1, v.Length - 1);
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
            }
        }

        const string _lblPos = "c:tickLblPos/@val";
        public eTickLabelPosition LabelPosition
        {
            get
            {
                return (eTickLabelPosition)Enum.Parse(typeof(eTickLabelPosition), GetXmlNodeString(_lblPos), true);
            }
            set
            {
                string lp = value.ToString();
                SetXmlNodeString(_lblPos, lp.Substring(0, 1).ToLower() + lp.Substring(1, lp.Length - 1));
            }
        }
        ExcelDrawingFill _fill = null;
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

        public bool Deleted 
        {
            get
            {
                return GetXmlNodeBool("c:delete/@val", false);
            }
            set
            {
                SetXmlNodeBool("c:delete/@val", value, false);
            }
        }
        const string _ticLblPos_Path = "c:tickLblPos/@val";
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
                v=v.Substring(0, 1).ToLower() + v.Substring(1, v.Length - 1);
                SetXmlNodeString(_ticLblPos_Path,v);
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
        /// <summary>
        /// Major unit for the axis.
        /// Null is automatic
        /// </summary>
        public double? MajorUnit
        {
            get
            {
                return GetXmlNodeDoubleNull(_majorUnitPath);
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_majorUnitPath);
                }
                else
                {
                    SetXmlNodeString(_majorUnitPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        const string _minorUnitPath = "c:minorUnit/@val";
        /// <summary>
        /// Minor unit for the axis.
        /// Null is automatic
        /// </summary>
        public double? MinorUnit
        {
            get
            {
                return GetXmlNodeDoubleNull(_minorUnitPath);
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_minorUnitPath);
                }
                else
                {
                    SetXmlNodeString(_minorUnitPath, ((double)value).ToString(CultureInfo.InvariantCulture));
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
        #endregion
    }
}
