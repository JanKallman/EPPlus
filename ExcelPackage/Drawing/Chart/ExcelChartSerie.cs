/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * EPPlus is a fork of the ExcelPackage project
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
using System.Collections;

namespace OfficeOpenXml.Drawing.Chart
{
   /// <summary>
   /// A chart serie
   /// </summary>
    public class ExcelChartSerie : XmlHelper
   {
       protected ExcelChartSeries _chartSeries;
       protected XmlNode _node;
       protected XmlNamespaceManager _ns;
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chartSeries">Parent collection</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
        public ExcelChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node)
           : base(ns,node)
       {
           _chartSeries = chartSeries;
           _node=node;
           _ns=ns;
           if (chartSeries.Chart.ChartType == eChartType.XYScatter ||
               chartSeries.Chart.ChartType == eChartType.XYScatterLines ||
               chartSeries.Chart.ChartType == eChartType.XYScatterLinesNoMarkers ||
               chartSeries.Chart.ChartType == eChartType.XYScatterSmooth ||
               chartSeries.Chart.ChartType == eChartType.XYScatterSmoothNoMarkers)
           {
               _seriesTopPath = "c:yVal";
               _xSeriesTopPath = "c:xVal";
               _seriesPath = string.Format(_seriesPath, _seriesTopPath);
               _xSeriesPath = string.Format(_xSeriesPath, _xSeriesTopPath);               
           }
           else
           {
               _seriesTopPath = "c:val";
               _xSeriesTopPath = "c:cat";
               _seriesPath = string.Format(_seriesPath, _seriesTopPath);
               _xSeriesPath = string.Format(_xSeriesPath, _xSeriesTopPath);
           }
       }
       const string headerPath="c:tx/c:v";
       /// <summary>
       /// Header for the serie.
       /// </summary>
       public string Header 
       {
           get
           {
                return GetXmlNode(headerPath);
            }
            set
            {
                //Where need this one 
                CreateNode(headerPath);
                SetXmlNode(headerPath, value);            
            }
        }

        string _seriesTopPath;
        string _seriesPath = "{0}/c:numRef/c:f";       
       /// <summary>
       /// Set this to a valid address or the drawing will be invalid.
       /// </summary>
       public string Series
       {
           get
           {
               return GetXmlNode(_seriesPath);
           }
           set
           {
               if (_chartSeries.Chart.ChartType == eChartType.Bubble)
               {
                   throw(new Exception("Bubble charts is not supported yet"));
               }
               CreateNode(_seriesPath,true);
               SetXmlNode(_seriesPath, ExcelCellBase.GetFullAddress(_chartSeries.Chart.WorkSheet.Name, value));
               
               XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:numRef/c:numCache",_seriesTopPath), _ns);
               if (cache != null)
               {
                   cache.ParentNode.RemoveChild(cache);
               }

               XmlNode lit = TopNode.SelectSingleNode(string.Format("{0}/c:numLit",_seriesTopPath), _ns);
               if (lit != null)
               {
                   lit.ParentNode.RemoveChild(lit);
               }
           }

       }
       string _xSeriesTopPath;
       string _xSeriesPath = "{0}/c:numRef/c:f";
       /// <summary>
       /// Set an address for the horisontal labels
       /// </summary>
       public string XSeries
       {
           get
           {
               return GetXmlNode(_xSeriesPath);
           }
           set
           {
               XmlNode node = TopNode.SelectSingleNode(_xSeriesTopPath, NameSpaceManager);
               if(node==null)
               {
                   node = TopNode.OwnerDocument.CreateElement(_xSeriesTopPath, ExcelPackage.schemaChart);
                   InserAfter(TopNode, "c:dLbls,c:tx,c:order", node);
               }
               SetXmlNode(_xSeriesPath, ExcelCellBase.GetFullAddress(_chartSeries.Chart.WorkSheet.Name, value));

               XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:numRef/c:numCache",_xSeriesTopPath), _ns);
               if (cache != null)
               {
                   cache.ParentNode.RemoveChild(cache);
               }

               XmlNode lit = TopNode.SelectSingleNode(string.Format("{0}/c:numLit",_xSeriesTopPath), _ns);
               if (lit != null)
               {
                   lit.ParentNode.RemoveChild(lit);
               }
           }
       }
   }
}
