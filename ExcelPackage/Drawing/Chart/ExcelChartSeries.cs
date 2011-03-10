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
using System.Collections;

namespace OfficeOpenXml.Drawing.Chart
{
   public sealed class ExcelChartSeries : XmlHelper, IEnumerable
    {
       List<ExcelChartSerie> _list=new List<ExcelChartSerie>();
       ExcelChart _chart;
       XmlNode _node;
       XmlNamespaceManager _ns;
       internal ExcelChartSeries (ExcelChart chart, XmlNamespaceManager ns, XmlNode node)
           : base(ns,node)
       {
           _ns = ns;
           _chart=chart;
           _node=node;
           SchemaNodeOrder = new string[] { "view3D", "plotArea", "barDir", "grouping", "scatterStyle", "varyColors", "ser", "explosion", "dLbls", "firstSliceAng", "holeSize", "shape", "legend", "axId" };
           foreach(XmlNode n in node.SelectNodes("c:ser",ns))
           {
               ExcelChartSerie s;
               if (chart.ChartNode.LocalName == "scatterChart")
               {
                   s= new ExcelScatterChartSerie(this, ns, n);
               }
               else if (chart.ChartNode.LocalName == "pieChart" ||
                        chart.ChartNode.LocalName == "ofPieChart" ||
                        chart.ChartNode.LocalName == "pie3DChart" ||
                        chart.ChartNode.LocalName == "doughnutChart")                                                                       
               {
                   s = new ExcelPieChartSerie(this, ns, n);
               }
               else
               {
                   s = new ExcelChartSerie(this, ns, n);
               }
               _list.Add(s);
           }
       }

       #region IEnumerable Members

       public IEnumerator GetEnumerator()
       {
           return (_list.GetEnumerator());
       }
       /// <summary>
       /// Returns the serie at the specified position.  
       /// </summary>
       /// <param name="PositionID">The position of the series.</param>
       /// <returns></returns>
       public ExcelChartSerie this[int PositionID]
       {
           get
           {
               return (_list[PositionID]);
           }
       }
       public int Count
       {
           get
           {
               return _list.Count;
           }
       }
       public void Delete(int PositionID)
       {
           ExcelChartSerie ser = _list[PositionID];
           ser.TopNode.ParentNode.RemoveChild(ser.TopNode);
           _list.RemoveAt(PositionID);
       }
       #endregion
       public ExcelChart Chart
       {
           get
           {
               return _chart;
           }
       }
       #region "Add Series"

       public ExcelChartSerie Add(ExcelRangeBase Serie, ExcelRangeBase XSerie)
       {
           return AddSeries(Serie.FullAddressAbsolute, XSerie.FullAddressAbsolute);
       }
       public ExcelChartSerie Add(string SerieAddress, string XSerieAddress)
       {
            return AddSeries(SerieAddress, XSerieAddress);
       }
       private ExcelChartSerie AddSeries(string SeriesAddress, string XSeriesAddress)
        {
               XmlElement ser = _node.OwnerDocument.CreateElement("ser", ExcelPackage.schemaChart);
               XmlNodeList node = _node.SelectNodes("c:ser", _ns);
               if (node.Count > 0)
               {
                   _node.InsertAfter(ser, node[node.Count-1]);
               }
               else
               {                   
                   InserAfter(_node, "c:grouping,c:barDir,c:scatterStyle,c:varyColors", ser);
                }
               int idx = FindIndex();
               ser.InnerXml = string.Format("<c:idx val=\"{1}\" /><c:order val=\"{1}\" /><c:tx><c:strRef><c:f></c:f><c:strCache><c:ptCount val=\"1\" /></c:strCache></c:strRef></c:tx>{5}{0}{2}{3}{4}", AddExplosion(Chart.ChartType), idx, AddScatterPoint(Chart.ChartType), AddAxisNodes(Chart.ChartType), AddSmooth(Chart.ChartType), AddMarker(Chart.ChartType));
               ExcelChartSerie serie;
               switch (Chart.ChartType)
               {
                   case eChartType.XYScatter:
                   case eChartType.XYScatterLines:
                   case eChartType.XYScatterLinesNoMarkers:
                   case eChartType.XYScatterSmooth:
                   case eChartType.XYScatterSmoothNoMarkers:
                       serie = new ExcelScatterChartSerie(this, NameSpaceManager, ser);
                       break;
                   case eChartType.Pie:
                   case eChartType.Pie3D:
                   case eChartType.PieExploded:
                   case eChartType.PieExploded3D:
                   case eChartType.PieOfPie:
                   case eChartType.Doughnut:
                   case eChartType.DoughnutExploded:
                   case eChartType.BarOfPie:
                       serie = new ExcelPieChartSerie(this, NameSpaceManager, ser);
                       break;
                   default:
                       serie = new ExcelChartSerie(this, NameSpaceManager, ser);
                       break;
               }
               serie.Series = SeriesAddress;
               serie.XSeries = XSeriesAddress;
               _list.Add(serie);
               return serie;
        }

       private int FindIndex()
       {
           int ret = 0, newID=0;
           if (_chart.PlotArea.ChartTypes.Count > 1)
           {
               bool chartAdded = false;
               foreach (var chart in _chart.PlotArea.ChartTypes)
               {
                   if (newID>0)
                   {
                       foreach (ExcelChartSerie serie in chart.Series)
                       {
                           serie.SetID((++newID).ToString());
                       }
                   }
                   else
                   {
                       if (chart == _chart)
                       {
                           ret += _list.Count + 1;
                           newID=ret;
                       }
                       else
                       {
                           ret += chart.Series.Count;
                       }
                   }
               }
               return ret-1;
           }
           else
           {
               return _list.Count;
           }
       }
       #endregion
       #region "Xml init Functions"
       private string AddMarker(eChartType chartType)
       {
           if (chartType == eChartType.Line ||
               chartType == eChartType.LineStacked ||
               chartType == eChartType.LineStacked100 ||
               chartType == eChartType.XYScatterLines ||
               chartType == eChartType.XYScatterSmooth ||
               chartType == eChartType.XYScatterLinesNoMarkers ||
               chartType == eChartType.XYScatterSmoothNoMarkers)
           {
               return "<c:marker><c:symbol val=\"none\" /></c:marker>";
           }
           else
           {
               return "";
           }
       }
       private string AddScatterPoint(eChartType chartType)
       {
           if (chartType == eChartType.XYScatter)
           {
               return "<c:spPr><a:ln w=\"28575\"><a:noFill /></a:ln></c:spPr>";
           }
           else
           {
               return "";
           }
       }
       private string AddAxisNodes(eChartType chartType)
       {
           if ( chartType == eChartType.XYScatter ||
                chartType == eChartType.XYScatterLines ||
                chartType == eChartType.XYScatterLinesNoMarkers ||
                chartType == eChartType.XYScatterSmooth ||
                chartType == eChartType.XYScatterSmoothNoMarkers)
           {
               return "<c:xVal /><c:yVal />";
           }
           else
           {
               return "<c:val />";
           }
       }

       private string AddExplosion(eChartType chartType)
       {
           if (chartType == eChartType.PieExploded3D ||
              chartType == eChartType.PieExploded ||
               chartType == eChartType.DoughnutExploded)
           {
               return "<c:explosion val=\"25\" />"; //Default 25;
           }
           else
           {
               return "";
           }
       }
       private string AddSmooth(eChartType chartType)
       {
           if (chartType == eChartType.XYScatterSmooth ||
              chartType == eChartType.XYScatterSmoothNoMarkers)
           {
               return "<c:smooth val=\"1\" />"; //Default 25;
           }
           else
           {
               return "";
           }
       }
        #endregion
    }
}
