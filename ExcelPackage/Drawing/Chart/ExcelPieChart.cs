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

    public class ExcelPieChart : ExcelChart
    {
        internal ExcelPieChart(ExcelDrawings drawings, XmlNode node, eChartType type) :
            base(drawings, node, type)
        {
        }
        internal ExcelPieChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart) :
            base(drawings, node, type, topChart)
        {
            //_varyColorsPath = string.Format(_varyColorsPath, GetChartNodeText());
        }

        public ExcelPieChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, PackagePart part, XmlDocument chartXml, XmlNode chartNode) :
           base(drawings, node, uriChart, part, chartXml, chartNode)
        {
        }

        public ExcelPieChart(ExcelChart topChart, XmlNode chartNode) :
            base(topChart, chartNode)
        {
        }
        ExcelChartDataLabel _DataLabel = null;
        private ExcelDrawings drawings;
        private XmlNode node;
        private ExcelChart topChart;
        public ExcelChartDataLabel DataLabel
        {
            get
            {
                if (_DataLabel == null)
                {
                    _DataLabel = new ExcelChartDataLabel(NameSpaceManager, ChartNode);
                }
                return _DataLabel;
            }
        }
        internal override eChartType GetChartType(string name)
        {
            if (name == "pieChart")
            {
                if (((ExcelPieChartSerie)Series[0]).Explosion>0)
                {
                    return eChartType.PieExploded;
                }
                else
                {
                    return eChartType.Pie;
                }
            }
            else if (name == "pie3DChart")
            {
                if (((ExcelPieChartSerie)Series[0]).Explosion > 0)
                {
                    return eChartType.PieExploded3D;
                }
                else
                {
                    return eChartType.Pie3D;
                }
            }
            return base.GetChartType(name);
        }
    }
}
