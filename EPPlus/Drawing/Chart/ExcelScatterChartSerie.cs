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
 * Jan Källman		Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using System.Drawing;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A serie for a scatter chart
    /// </summary>
    public sealed class ExcelScatterChartSerie : ExcelChartSerie
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chartSeries">Parent collection</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
        /// <param name="isPivot">Is pivotchart</param>
        internal ExcelScatterChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node, bool isPivot) :
            base(chartSeries, ns, node, isPivot)
        {
            if (chartSeries.Chart.ChartType == eChartType.XYScatterLines ||
                chartSeries.Chart.ChartType == eChartType.XYScatterSmooth)
            {
                Marker = eMarkerStyle.Square;
            }

            if (chartSeries.Chart.ChartType == eChartType.XYScatterSmooth ||
                chartSeries.Chart.ChartType == eChartType.XYScatterSmoothNoMarkers)
            {
                Smooth = 1;
            }
            else if (chartSeries.Chart.ChartType == eChartType.XYScatterLines || chartSeries.Chart.ChartType == eChartType.XYScatterLinesNoMarkers || chartSeries.Chart.ChartType == eChartType.XYScatter)

            {
                Smooth = 0;
            }
        }
        ExcelChartSerieDataLabel _DataLabel = null;
        /// <summary>
        /// Datalabel
        /// </summary>
        public ExcelChartSerieDataLabel DataLabel
        {
            get
            {
                if (_DataLabel == null)
                {
                    _DataLabel = new ExcelChartSerieDataLabel(_ns, _node);
                }
                return _DataLabel;
            }
        }
        const string smoothPath = "c:smooth/@val";
        /// <summary>
        /// Smooth for scattercharts
        /// </summary>
        public int Smooth
        {
            get
            {
                return GetXmlNodeInt(smoothPath);
            }
            internal set
            {
                SetXmlNodeString(smoothPath, value.ToString());
            }
        }
        const string markerPath = "c:marker/c:symbol/@val";
        /// <summary>
        /// Marker symbol 
        /// </summary>
        public eMarkerStyle Marker
        {
            get
            {
                string marker = GetXmlNodeString(markerPath);
                if (marker == "")
                {
                    return eMarkerStyle.None;
                }
                else
                {
                    return (eMarkerStyle)Enum.Parse(typeof(eMarkerStyle), marker, true);
                }
            }
            set
            {
                /* setting MarkerStyle seems to be working, so no need to throw an exception in this case
                 if (_chartSeries.Chart.ChartType == eChartType.XYScatterLinesNoMarkers ||
                    _chartSeries.Chart.ChartType == eChartType.XYScatterSmoothNoMarkers)
                {
                    throw (new InvalidOperationException("Can't set marker style for this charttype."));
                }*/
                SetXmlNodeString(markerPath, value.ToString().ToLower(CultureInfo.InvariantCulture));
            }
        }

        //new properties for excel scatter-plots: LineColor, MarkerSize, MarkerColor, LineWidth and MarkerLineColor
        //implemented according to https://epplus.codeplex.com/discussions/287917
        string LINECOLOR_PATH = "c:spPr/a:ln/a:solidFill/a:srgbClr/@val";
        /// <summary>
        /// Line color.
        /// </summary>
        ///
        /// <value>
        /// The color of the line.
        /// </value>
        public Color LineColor
        {
            get
            {
                string color = GetXmlNodeString(LINECOLOR_PATH);
                if (color == "")
                {
                    return Color.Black;
                }
                else
                {
                    Color c = Color.FromArgb(Convert.ToInt32(color, 16));
                    int a = getAlphaChannel(LINECOLOR_PATH);
                    if (a != 255)
                    {
                        c = Color.FromArgb(a, c);
                    }
                    return c;
                }
            }
            set
            {
                SetXmlNodeString(LINECOLOR_PATH, value.ToArgb().ToString("X8").Substring(2), true);
                setAlphaChannel(value, LINECOLOR_PATH);
            }
        }
        string MARKERSIZE_PATH = "c:marker/c:size/@val";
        /// <summary>
        /// Gets or sets the size of the marker.
        /// </summary>
        ///
        /// <remarks>
        /// value between 2 and 72.
        /// </remarks>
        ///
        /// <value>
        /// The size of the marker.
        /// </value>
        public int MarkerSize
        {
            get
            {
                string size = GetXmlNodeString(MARKERSIZE_PATH);
                if (size == "")
                {
                    return 5;
                }
                else
                {
                    return Int32.Parse(GetXmlNodeString(MARKERSIZE_PATH));
                }
            }
            set
            {
                int size = value;
                size = Math.Max(2, size);
                size = Math.Min(72, size);
                SetXmlNodeString(MARKERSIZE_PATH, size.ToString(), true);
            }
        }
        string MARKERCOLOR_PATH = "c:marker/c:spPr/a:solidFill/a:srgbClr/@val";
        /// <summary>
        /// Marker color.
        /// </summary>
        ///
        /// <value>
        /// The color of the Marker.
        /// </value>
        public Color MarkerColor
        {
            get
            {
                string color = GetXmlNodeString(MARKERCOLOR_PATH);
                if (color == "")
                {
                    return Color.Black;
                }
                else
                {
                    Color c = Color.FromArgb(Convert.ToInt32(color, 16));
                    int a = getAlphaChannel(MARKERCOLOR_PATH);
                    if (a != 255)
                    {
                        c = Color.FromArgb(a, c);
                    }
                    return c;
                }
            }
            set
            {
                SetXmlNodeString(MARKERCOLOR_PATH, value.ToArgb().ToString("X8").Substring(2), true); //.Substring(2) => cut alpha value
                setAlphaChannel(value, MARKERCOLOR_PATH);
            }
        }

        string LINEWIDTH_PATH = "c:spPr/a:ln/@w";
        /// <summary>
        /// Gets or sets the width of the line in pt.
        /// </summary>
        ///
        /// <value>
        /// The width of the line.
        /// </value>
        public double LineWidth
        {
            get
            {
                string size = GetXmlNodeString(LINEWIDTH_PATH);
                if (size == "")
                {
                    return 2.25;
                }
                else
                {
                    return double.Parse(GetXmlNodeString(LINEWIDTH_PATH)) / 12700;
                }
            }
            set
            {
                SetXmlNodeString(LINEWIDTH_PATH, (( int )(12700 * value)).ToString(), true);
            }
        }
        //marker line color
        string MARKERLINECOLOR_PATH = "c:marker/c:spPr/a:ln/a:solidFill/a:srgbClr/@val";
        /// <summary>
        /// Marker Line color.
        /// (not to be confused with LineColor)
        /// </summary>
        ///
        /// <value>
        /// The color of the Marker line.
        /// </value>
        public Color MarkerLineColor
        {
            get
            {
                string color = GetXmlNodeString(MARKERLINECOLOR_PATH);
                if (color == "")
                {
                    return Color.Black;
                }
                else
                {
                    Color c = Color.FromArgb(Convert.ToInt32(color, 16));
                    int a = getAlphaChannel(MARKERLINECOLOR_PATH);
                    if (a != 255)
                    {
                        c = Color.FromArgb(a, c);
                    }
                    return c;
                }
            }
            set
            {
                SetXmlNodeString(MARKERLINECOLOR_PATH, value.ToArgb().ToString("X8").Substring(2), true);
                setAlphaChannel(value, MARKERLINECOLOR_PATH);
            }
        }


        /// <summary>
        /// write alpha value (if Color.A != 255)
        /// </summary>
        /// <param name="c">Color</param>
        /// <param name="xPath">where to write</param>
        /// <remarks>
        /// alpha-values may only written to color-nodes
        /// eg: a:prstClr (preset), a:hslClr (hsl), a:schemeClr (schema), a:sysClr (system), a:scrgbClr (rgb percent) or a:srgbClr (rgb hex)
        ///     .../a:prstClr/a:alpha/@val
        /// </remarks>
        private void setAlphaChannel(Color c, string xPath)
        {
            //check 4 Alpha-values
            if (c.A != 255)
            { //opaque color => alpha == 255 //source: https://msdn.microsoft.com/en-us/library/1hstcth9%28v=vs.110%29.aspx
                //check path
                string s = xPath4Alpha(xPath);
                if (s.Length > 0)
                {
                    string alpha = ((c.A == 0) ? 0 : (100 - c.A) * 1000).ToString(); //note: excel writes 100% transparency (alpha=0) as "0" and not as "100000"
                    SetXmlNodeString(s, alpha, true);
                }
            }
        }
        /// <summary>
        /// read AlphaChannel from a:solidFill
        /// </summary>
        /// <param name="xPath"></param>
        /// <returns>alpha or 255 if their is no such node</returns>
        private int getAlphaChannel(string xPath)
        {
            int r = 255;
            string s = xPath4Alpha(xPath);
            if (s.Length > 0)
            {
                int i = 0;
                if (int.TryParse(GetXmlNodeString(s), out i))
                {
                    r = (i == 0) ? 0 : 100 - (i / 1000);
                }
            }
            return r;
        }
        /// <summary>
        /// creates xPath to alpha attribute for a color 
        /// eg: a:prstClr/a:alpha/@val
        /// </summary>
        /// <param name="xPath">xPath to color node</param>
        /// <returns></returns>
        private string xPath4Alpha(string xPath)
        {
            string s = string.Empty;
            if (xPath.EndsWith("@val"))
            {
                xPath = xPath.Substring(0, xPath.IndexOf("@val"));
            }
            if (xPath.EndsWith("/"))
            { //cut tailing slash
                xPath = xPath.Substring(0, xPath.Length - 1);
            }
            //parent node must be a color node/definition
            List<string> colorDefs = new List<string>() { "a:prstClr", "a:hslClr", "a:schemeClr", "a:sysClr", "a:scrgbClr", "a:srgbClr" };
            if (colorDefs.Find(cd => xPath.EndsWith(cd, StringComparison.InvariantCulture)) != null)
            {
                s = xPath + "/a:alpha/@val";
            }
            else
            {
                System.Diagnostics.Debug.Assert(false);
                throw new InvalidOperationException("alpha-values can only set to Colors");
            }
            return s;
        }


    }
}
