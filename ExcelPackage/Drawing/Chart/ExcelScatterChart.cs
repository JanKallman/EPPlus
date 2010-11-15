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
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO.Packaging;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Provides access to scatter chart specific properties
    /// </summary>
    public sealed class ExcelScatterChart : ExcelChart
    {
        internal ExcelScatterChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart) :
            base(drawings, node, type, topChart)
        {
            SetTypeProperties();
        }

        internal ExcelScatterChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, PackagePart part, XmlDocument chartXml, XmlNode chartNode) :
            base(drawings, node, uriChart, part, chartXml, chartNode)
        {
            SetTypeProperties();
        }

        internal ExcelScatterChart(ExcelChart topChart, XmlNode chartNode) : 
            base(topChart, chartNode)
        {
            SetTypeProperties();
        }
        private void SetTypeProperties()
        {
           /***** ScatterStyle *****/
           if(ChartType == eChartType.XYScatter ||
              ChartType == eChartType.XYScatterLines ||
              ChartType == eChartType.XYScatterLinesNoMarkers)
           {
               ScatterStyle = eScatterStyle.LineMarker;
          }
           else if (
              ChartType == eChartType.XYScatterSmooth ||
              ChartType == eChartType.XYScatterSmoothNoMarkers) 
           {
               ScatterStyle = eScatterStyle.SmoothMarker;
           }
        }
        #region "Grouping Enum Translation"
        string _scatterTypePath = "c:scatterStyle/@val";
        private eScatterStyle GetScatterEnum(string text)
        {
            switch (text)
            {
                case "smoothMarker":
                    return eScatterStyle.SmoothMarker;
                default:
                    return eScatterStyle.LineMarker;
            }
        }

        private string GetScatterText(eScatterStyle shatterStyle)
        {
            switch (shatterStyle)
            {
                case eScatterStyle.SmoothMarker:
                    return "smoothMarker";
                default:
                    return "lineMarker";
            }
        }
        #endregion
        /// <summary>
        /// If the scatter has LineMarkers or SmoothMarkers
        /// </summary>
        public eScatterStyle ScatterStyle
        {
            get
            {
                return GetScatterEnum(_chartXmlHelper.GetXmlNodeString(_scatterTypePath));
            }
            internal set
            {
                _chartXmlHelper.CreateNode(_scatterTypePath, true);
                _chartXmlHelper.SetXmlNodeString(_scatterTypePath, GetScatterText(value));
            }
        }
        string MARKER_PATH = "c:marker/@val";
        /// <summary>
        /// If the series has markers
        /// </summary>
        public bool Marker
        {
            get
            {
                return GetXmlNodeBool(MARKER_PATH, false);
            }
            set
            {
                SetXmlNodeBool(MARKER_PATH, value, false);
            }
        }
        internal override eChartType GetChartType(string name)
        {
            if (name == "scatterChart")
            {
                if (ScatterStyle==eScatterStyle.LineMarker)
                {
                    if (((ExcelScatterChartSerie)Series[0]).Marker == eMarkerStyle.None)
                    {
                        return eChartType.XYScatterLinesNoMarkers;
                    }
                    else
                    {
                        if(ExistNode("c:ser/c:spPr/a:ln/noFill"))
                        {
                            return eChartType.XYScatter;
                        }
                        else
                        {
                            return eChartType.XYScatterLines;
                        }
                    }
                }
                else if (ScatterStyle == eScatterStyle.SmoothMarker)
                {
                    if (((ExcelScatterChartSerie)Series[0]).Marker == eMarkerStyle.None)
                    {
                        return eChartType.XYScatterSmoothNoMarkers;
                    }
                    else
                    {
                        return eChartType.XYScatterSmooth;
                    }
                }
            }
            return base.GetChartType(name);
        }

    }
}
