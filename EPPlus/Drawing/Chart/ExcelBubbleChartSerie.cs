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
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A serie for a scatter chart
    /// </summary>
    public sealed class ExcelBubbleChartSerie : ExcelChartSerie
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chartSeries">Parent collection</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
        /// <param name="isPivot">Is pivotchart</param>
        internal ExcelBubbleChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node, bool isPivot) :
            base(chartSeries, ns, node, isPivot)
        {
            if (chartSeries.Chart.ChartType == eChartType.Bubble3DEffect)
            {
                
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
        const string BUBBLE3D_PATH = "c:bubble3D/@val";
        internal bool Bubble3D
        {
            get
            {
                return GetXmlNodeBool(BUBBLE3D_PATH, true);
            }
            set
            {
                SetXmlNodeBool(BUBBLE3D_PATH, value);    
            }
        }
        public override string Series
        {
            get
            {
                return base.Series;
            }
            set
            {
                base.Series = value;
                if(string.IsNullOrEmpty(BubbleSize))
                {
                    GenerateLit();
                }
            }
        }
        const string BUBBLESIZE_TOPPATH = "c:bubbleSize";
        const string BUBBLESIZE_PATH = BUBBLESIZE_TOPPATH + "/c:numRef/c:f";
        public string BubbleSize
        {
            get
            {
                return GetXmlNodeString(BUBBLESIZE_PATH);
            }
            set
            {
                if(string.IsNullOrEmpty(value))
                {
                    GenerateLit();
                }
                else
                {
                    SetXmlNodeString(BUBBLESIZE_PATH, ExcelCellBase.GetFullAddress(_chartSeries.Chart.WorkSheet.Name, value));
                
                    XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:numCache", BUBBLESIZE_PATH), _ns);
                    if (cache != null)
                    {
                        cache.ParentNode.RemoveChild(cache);
                    }

                    DeleteNode(string.Format("{0}/c:numLit", BUBBLESIZE_TOPPATH));
                    //XmlNode lit = TopNode.SelectSingleNode(string.Format("{0}/c:numLit", _xSeriesTopPath), _ns);
                    //if (lit != null)
                    //{
                    //    lit.ParentNode.RemoveChild(lit);
                    //}
                }
            }
        }

        internal void GenerateLit()
        {
            var s = new ExcelAddress(Series);
            var ix = 0;
            var sb = new StringBuilder();
            for (int row = s._fromRow; row <= s._toRow; row++)
            {
                for (int c = s._fromCol; c <= s._toCol; c++)
                {
                    sb.AppendFormat("<c:pt idx=\"{0}\"><c:v>1</c:v></c:pt>", ix++);
                }
            }
            CreateNode(BUBBLESIZE_TOPPATH + "/c:numLit", true);
            XmlNode lit = TopNode.SelectSingleNode(string.Format("{0}/c:numLit", BUBBLESIZE_TOPPATH), _ns);
            lit.InnerXml = string.Format("<c:formatCode>General</c:formatCode><c:ptCount val=\"{0}\"/>{1}", ix, sb.ToString());
        }
    }
}
