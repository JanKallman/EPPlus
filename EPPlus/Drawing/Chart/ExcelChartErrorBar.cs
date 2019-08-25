/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
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
 * ******************************************************************************
 * Kris Wragg		            Initial Release		            2019-08-25
 *******************************************************************************/

using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    public class ExcelChartErrorBar : XmlHelper
    {
        internal ExcelChartSerie _chartSerie;
        protected XmlNode _node;
        protected XmlNamespaceManager _ns;

        const string BARDIRPATH = "c:errDir/@val";
        const string BARTYPEPATH = "c:errBarType/@val";
        const string VALTYPEPATH = "c:errValType/@val";
        const string NOENDCAPVALUEPATH = "c:noEndCap/@val";

        const string _errorBarValuePath = "c:val/@val";
        const string _minusErrorPath = "c:minus/c:numRef/c:f";
        const string _plusErrorPath = "c:plus/c:numRef/c:f";
        const string _minusErrorCachePath = "c:minus/c:numRef/c:numCache";
        const string _plusErrorCachePath = "c:plus/c:numRef/c:numCache";
        const string _minusErrorLitPath = "c:minus/c:numRef/c:numLit";
        const string _plusErrorLitPath = "c:plus/c:numRef/c:numLit";

        internal ExcelChartErrorBar(ExcelChartSerie chartSerie, XmlNamespaceManager ns, XmlNode node)
            : base(ns, node)
        {
            _chartSerie = chartSerie;
            _node = node;
            _ns = ns;

            SchemaNodeOrder = new string[] { "errDir", "errBarType", "errValType", "noEndCap", "plus", "minus", "val", "spPr" };
        }

        internal ExcelChartErrorBar(ExcelChartSerie chartSerie, XmlNamespaceManager ns, XmlNode node, eErrorBarDirection direction)
            : this(chartSerie, ns, node)
        {
            Direction = direction;
        }

        /// <summary>
        /// The direction of the error bar - X or Y.
        /// </summary>
        public eErrorBarDirection Direction
        {
            get
            {
                switch (GetXmlNodeString(BARDIRPATH).ToLower(CultureInfo.InvariantCulture))
                {
                    case "x":
                        return eErrorBarDirection.X;
                    case "y":
                        return eErrorBarDirection.Y;
                    default:
                        return eErrorBarDirection.X;
                }
            }

            internal set
            {
                switch (value)
                {
                    case eErrorBarDirection.X:
                        SetXmlNodeString(BARDIRPATH, "x");
                        break;
                    case eErrorBarDirection.Y:
                        SetXmlNodeString(BARDIRPATH, "y");
                        break;
                    default:
                        SetXmlNodeString(BARDIRPATH, "x");
                        break;
                }
            }
        }

        /// <summary>
        /// The type of the error bars - positive, negative, or both.
        /// </summary>
        public eErrorBarType Type
        {
            get
            {
                switch (GetXmlNodeString(BARTYPEPATH).ToLower(CultureInfo.InvariantCulture))
                {
                    case "both":
                        return eErrorBarType.Both;
                    case "minus":
                        return eErrorBarType.Minus;
                    case "plus":
                        return eErrorBarType.Plus;
                    default:
                        return eErrorBarType.Both;
                }
            }

            set
            {
                switch (value)
                {
                    case eErrorBarType.Both:
                        SetXmlNodeString(BARTYPEPATH, "both");
                        break;
                    case eErrorBarType.Minus:
                        SetXmlNodeString(BARTYPEPATH, "minus");
                        break;
                    case eErrorBarType.Plus:
                        SetXmlNodeString(BARTYPEPATH, "plus");
                        break;
                    default:
                        SetXmlNodeString(BARTYPEPATH, "both");
                        break;
                }
            }
        }

        /// <summary>
        /// The type of values used to determine the length of the error bars.
        /// </summary>
        public eErrorBarValueType ValueType
        {
            get
            {
                switch (GetXmlNodeString(VALTYPEPATH).ToLower(CultureInfo.InvariantCulture))
                {
                    case "cust":
                        return eErrorBarValueType.CustomErrorBars;
                    case "fixedVal":
                        return eErrorBarValueType.FixedValue;
                    case "percentage":
                        return eErrorBarValueType.Percentage;
                    case "stdDev":
                        return eErrorBarValueType.StandardDeviation;
                    case "stdErr":
                        return eErrorBarValueType.StandardError;
                    default:
                        return eErrorBarValueType.FixedValue;
                }
            }

            set
            {
                switch (value)
                {
                    case eErrorBarValueType.CustomErrorBars:
                        SetXmlNodeString(VALTYPEPATH, "cust");
                        break;
                    case eErrorBarValueType.FixedValue:
                        SetXmlNodeString(VALTYPEPATH, "fixedVal");
                        break;
                    case eErrorBarValueType.Percentage:
                        SetXmlNodeString(VALTYPEPATH, "percentage");
                        break;
                    case eErrorBarValueType.StandardDeviation:
                        SetXmlNodeString(VALTYPEPATH, "stdDev");
                        break;
                    case eErrorBarValueType.StandardError:
                        SetXmlNodeString(VALTYPEPATH, "stdErr");
                        break;
                    default:
                        SetXmlNodeString(VALTYPEPATH, "fixedVal");
                        break;
                }
            }
        }

        /// <summary>
        /// This element specifies whether an end cap is not drawn on the error bars.
        /// </summary>
        public bool NoEndCap
        {
            get
            {
                return GetXmlNodeBool(NOENDCAPVALUEPATH, true);
            }

            set
            {
                SetXmlNodeBool(NOENDCAPVALUEPATH, value, true);
            }
        }

        /// <summary>
        /// Address range used for Custom value type
        /// </summary>
        public string MinusAddress
        {
            get
            {
                return GetXmlNodeString(_minusErrorPath);
            }

            set
            {
                CreateNode(_minusErrorPath, true);
                SetXmlNodeString(_minusErrorPath, ExcelCellBase.GetFullAddress(_chartSerie._chartSeries.Chart.WorkSheet.Name, value));

                CleanupCacheAndLit(_minusErrorCachePath, _minusErrorLitPath);
             }
        }

        /// <summary>
        /// Address range used for Custom value type
        /// </summary>
        public string PlusAddress
        {
            get
            {
                return GetXmlNodeString(_plusErrorPath);
            }

            set
            {
                CreateNode(_plusErrorPath, true);
                SetXmlNodeString(_plusErrorPath, ExcelCellBase.GetFullAddress(_chartSerie._chartSeries.Chart.WorkSheet.Name, value));

                CleanupCacheAndLit(_plusErrorCachePath, _plusErrorLitPath);
            }
        }

        /// <summary>
        /// This element specifies a value which is used with the Error Bar Type to determine the length of the error bars.
        /// </summary>
        public double Value
        {
            get
            {
                double? value =  GetXmlNodeDoubleNull(_errorBarValuePath);

                switch(ValueType)
                {
                    case eErrorBarValueType.CustomErrorBars:
                        throw new Exception("Error bar value is not valid for Custom Error Bars, use PlusAddress and MinusAddress");
                    case eErrorBarValueType.StandardError:
                        throw new Exception("Error bar value is not valid for Standard Error");
                    case eErrorBarValueType.FixedValue:
                        value = 0.1;
                        break;
                    case eErrorBarValueType.Percentage:
                        value = 5;
                        break;
                    case eErrorBarValueType.StandardDeviation:
                        value = 1.0;
                        break;
                }

                return value.Value;
            }

            set
            {
                SetXmlNodeString(_errorBarValuePath, value.ToString(CultureInfo.InvariantCulture));
            }
        }

        ExcelDrawingBorder _fill = null;
        public ExcelDrawingBorder Line
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingBorder(NameSpaceManager, TopNode, "c:spPr/a:ln");
                }

                return _fill;
            }
        }

        private void CleanupCacheAndLit(string cachePath, string litPath)
        {
            XmlNode cache = TopNode.SelectSingleNode(cachePath, _ns);
            if (cache != null)
            {
                cache.ParentNode.RemoveChild(cache);
            }

            XmlNode lit = TopNode.SelectSingleNode(litPath, _ns);
            if (lit != null)
            {
                lit.ParentNode.RemoveChild(lit);
            }
        }
    }
}
