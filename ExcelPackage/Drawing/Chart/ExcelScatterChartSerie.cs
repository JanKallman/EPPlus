/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * EPPlus is a fork of the ExcelPackage project
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
 *******************************************************************************
 * Jan Källman		Added		30-DEC-2009
 *******************************************************************************/

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    public class ExcelScatterChartSerie : ExcelChartSerie
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chartSeries">Parent collection</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
        internal ExcelScatterChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node) :
            base(chartSeries, ns, node)
        {
            if (chartSeries.Chart.ChartType == eChartType.XYScatterLines ||
                chartSeries.Chart.ChartType == eChartType.XYScatterSmooth)
            {
                Marker = eMarkerStyle.Square;
            }
        }
        ExcelChartSerieDataLabel _DataLabel = null;
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
                SetXmlNode(smoothPath, value.ToString());
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
                string marker = GetXmlNode(markerPath);
                if (marker == "")
                {
                    return eMarkerStyle.Square;
                }
                else
                {
                    return (eMarkerStyle)Enum.Parse(typeof(eMarkerStyle), marker, true);
                }
            }
            internal set
            {
                SetXmlNode(markerPath, value.ToString().ToLower());
            }
        }        
       
    }
}
