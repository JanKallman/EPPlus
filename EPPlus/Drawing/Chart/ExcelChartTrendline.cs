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
 * ******************************************************************************
 * Jan Källman		Initial Release		        2011-05-25
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Globalization;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A collection of trendlines.
    /// </summary>
    public class ExcelChartTrendlineCollection : IEnumerable<ExcelChartTrendline>
    {
        List<ExcelChartTrendline> _list = new List<ExcelChartTrendline>();
        ExcelChartSerie _serie;
        internal ExcelChartTrendlineCollection(ExcelChartSerie serie)
        {
            _serie = serie;
            foreach (XmlNode node in _serie.TopNode.SelectNodes("c:trendline", _serie.NameSpaceManager))
            {
                _list.Add(new ExcelChartTrendline(_serie.NameSpaceManager, node));
            }
        }
        /// <summary>
        /// Add a new trendline
        /// </summary>
        /// <param name="Type"></param>
        /// <returns>The trendline</returns>
        public ExcelChartTrendline Add(eTrendLine Type)
        {
            if (_serie._chartSeries._chart.IsType3D() ||
                _serie._chartSeries._chart.IsTypePercentStacked() ||
                _serie._chartSeries._chart.IsTypeStacked() ||
                _serie._chartSeries._chart.IsTypePieDoughnut())
            {
                throw(new ArgumentException("Trendlines don't apply to 3d-charts, stacked charts, pie charts or doughnut charts"));
            }
            ExcelChartTrendline tl;
            XmlNode insertAfter;
            if (_list.Count > 0)
            {
                insertAfter = _list[_list.Count - 1].TopNode;
            }
            else
            {
                insertAfter = _serie.TopNode.SelectSingleNode("c:marker", _serie.NameSpaceManager);
                if (insertAfter == null)
                {
                    insertAfter = _serie.TopNode.SelectSingleNode("c:tx", _serie.NameSpaceManager);
                    if (insertAfter == null)
                    {
                        insertAfter = _serie.TopNode.SelectSingleNode("c:order", _serie.NameSpaceManager);
                    }
                }
            }
            var node=_serie.TopNode.OwnerDocument.CreateElement("c","trendline", ExcelPackage.schemaChart);
            _serie.TopNode.InsertAfter(node, insertAfter);

            tl = new ExcelChartTrendline(_serie.NameSpaceManager, node);
            tl.Type = Type;
            return tl;
        }
        IEnumerator<ExcelChartTrendline> IEnumerable<ExcelChartTrendline>.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
    }
    /// <summary>
    /// A trendline object
    /// </summary>
    public class ExcelChartTrendline : XmlHelper
    {
        internal ExcelChartTrendline(XmlNamespaceManager namespaceManager, XmlNode topNode) :
            base(namespaceManager,topNode)

        {
            SchemaNodeOrder = new string[] { "name", "trendlineType","order","period", "forward","backward","intercept", "dispRSqr", "dispEq", "trendlineLbl" };
        }
        const string TRENDLINEPATH = "c:trendlineType/@val";
        /// <summary>
        /// Type of Trendline
        /// </summary>
        public eTrendLine Type
        {
           get
           {
               switch (GetXmlNodeString(TRENDLINEPATH).ToLower())
               {
                   case "exp":
                       return eTrendLine.Exponential;
                   case "log":
                        return eTrendLine.Logarithmic;
                   case "poly":
                       return eTrendLine.Polynomial;
                   case "movingavg":
                       return eTrendLine.MovingAvgerage;
                   case "power":
                       return eTrendLine.Power;
                   default:
                       return eTrendLine.Linear;
               }
           }
           set
           {
                switch (value)
                {
                    case eTrendLine.Exponential:
                        SetXmlNodeString(TRENDLINEPATH, "exp");
                        break;
                    case eTrendLine.Logarithmic:
                        SetXmlNodeString(TRENDLINEPATH, "log");
                        break;
                    case eTrendLine.Polynomial:
                        SetXmlNodeString(TRENDLINEPATH, "poly");
                        Order = 2;
                        break;
                    case eTrendLine.MovingAvgerage:
                        SetXmlNodeString(TRENDLINEPATH, "movingAvg");
                        Period = 2;
                        break;
                    case eTrendLine.Power:
                        SetXmlNodeString(TRENDLINEPATH, "power");
                        break;
                    default: 
                        SetXmlNodeString(TRENDLINEPATH, "linear");
                        break;
                }
           }
        }
        const string NAMEPATH = "c:name";
        /// <summary>
        /// Name in the legend
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString(NAMEPATH);
            }
            set
            {
                SetXmlNodeString(NAMEPATH, value, true);
            }
        }
        const string ORDERPATH = "c:order/@val";
        /// <summary>
        /// Order for polynominal trendlines
        /// </summary>
        public decimal Order
        {
            get
            {
                return GetXmlNodeDecimal(ORDERPATH);
            }
            set
            {
                if (Type == eTrendLine.MovingAvgerage)
                {
                    throw (new ArgumentException("Can't set period for trendline type MovingAvgerage"));
                }
                DeleteAllNode(PERIODPATH);
                SetXmlNodeString(ORDERPATH, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        const string PERIODPATH = "c:period/@val";
        /// <summary>
        /// Period for monthly average trendlines
        /// </summary>
        public decimal Period
        {
            get
            {
                return GetXmlNodeDecimal(PERIODPATH);
            }
            set
            {
                if (Type == eTrendLine.Polynomial)
                {
                    throw (new ArgumentException("Can't set period for trendline type Polynomial"));
                }
                DeleteAllNode(ORDERPATH);
                SetXmlNodeString(PERIODPATH, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        const string FORWARDPATH = "c:forward/@val";
        /// <summary>
        /// Forcast forward periods
        /// </summary>
        public decimal Forward
        {
            get
            {
                return GetXmlNodeDecimal(FORWARDPATH);
            }
            set
            {
                SetXmlNodeString(FORWARDPATH, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        const string BACKWARDPATH = "c:backward/@val";
        /// <summary>
        /// Forcast backwards periods
        /// </summary>
        public decimal Backward
        {
            get
            {
                return GetXmlNodeDecimal(BACKWARDPATH);
            }
            set
            {
                SetXmlNodeString(BACKWARDPATH, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        const string INTERCEPTPATH = "c:intercept/@val";
        /// <summary>
        /// Specify the point where the trendline crosses the vertical axis
        /// </summary>
        public decimal Intercept
        {
            get
            {
                return GetXmlNodeDecimal(INTERCEPTPATH);
            }
            set
            {
                SetXmlNodeString(INTERCEPTPATH, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        const string DISPLAYRSQUAREDVALUEPATH = "c:dispRSqr/@val";
        /// <summary>
        /// Display the R-squared value for a trendline
        /// </summary>
        public bool DisplayRSquaredValue
        {
            get
            {
                return GetXmlNodeBool(DISPLAYRSQUAREDVALUEPATH, false);
            }
            set
            {
                SetXmlNodeBool(DISPLAYRSQUAREDVALUEPATH,value, false);
            }
        }
        const string DISPLAYEQUATIONPATH = "c:dispEq/@val";
        /// <summary>
        /// Display the trendline equation on the chart
        /// </summary>
        public bool DisplayEquation
        {
            get
            {
                return GetXmlNodeBool(DISPLAYEQUATIONPATH, false);
            }
            set
            {
                SetXmlNodeBool(DISPLAYEQUATIONPATH, value, false);
            }
        }
    }
}
