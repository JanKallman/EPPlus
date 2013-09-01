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
 *******************************************************************************
 * Jan Källman		Added		2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Collections;
using OfficeOpenXml.Table.PivotTable;
namespace OfficeOpenXml.Drawing.Chart
{
    public sealed class ExcelBubbleChartSeries : ExcelChartSeries
    {
        internal ExcelBubbleChartSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot)
           : base(chart,ns,node, isPivot)
        {
            //_chartSeries = new ExcelChartSeries(this, _drawings.NameSpaceManager, _chartNode, isPivot);
        }
        public ExcelChartSerie Add(ExcelRangeBase Serie, ExcelRangeBase XSerie, ExcelRangeBase BubbleSize)
        {
            return base.AddSeries(Serie.FullAddressAbsolute, XSerie.FullAddressAbsolute, BubbleSize.FullAddressAbsolute);
        }
        public ExcelChartSerie Add(string SerieAddress, string XSerieAddress, string BubbleSizeAddress)
        {
            return base.AddSeries(SerieAddress, XSerieAddress, BubbleSizeAddress);
        }
    }
    /// <summary>
   /// Collection class for chart series
   /// </summary>
    public class ExcelChartSeries : XmlHelper, IEnumerable
    {
       List<ExcelChartSerie> _list=new List<ExcelChartSerie>();
       internal ExcelChart _chart;
       XmlNode _node;
       XmlNamespaceManager _ns;
       internal ExcelChartSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot)
           : base(ns,node)
       {
           _ns = ns;
           _chart=chart;
           _node=node;
           _isPivot = isPivot;
           SchemaNodeOrder = new string[] { "view3D", "plotArea", "barDir", "grouping", "scatterStyle", "varyColors", "ser", "explosion", "dLbls", "firstSliceAng", "holeSize", "shape", "legend", "axId" };
           foreach(XmlNode n in node.SelectNodes("c:ser",ns))
           {
               ExcelChartSerie s;
               if (chart.ChartNode.LocalName == "scatterChart")
               {
                   s = new ExcelScatterChartSerie(this, ns, n, isPivot);
               }
               else if (chart.ChartNode.LocalName == "lineChart")
               {
                   s = new ExcelLineChartSerie(this, ns, n, isPivot);
               }
               else if (chart.ChartNode.LocalName == "pieChart" ||
                        chart.ChartNode.LocalName == "ofPieChart" ||
                        chart.ChartNode.LocalName == "pie3DChart" ||
                        chart.ChartNode.LocalName == "doughnutChart")                                                                       
               {
                   s = new ExcelPieChartSerie(this, ns, n, isPivot);
               }
               else
               {
                   s = new ExcelChartSerie(this, ns, n, isPivot);
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
       /// <summary>
       /// Delete the chart at the specific position
       /// </summary>
       /// <param name="PositionID">Zero based</param>
       public void Delete(int PositionID)
       {
           ExcelChartSerie ser = _list[PositionID];
           ser.TopNode.ParentNode.RemoveChild(ser.TopNode);
           _list.RemoveAt(PositionID);
       }
       #endregion
       /// <summary>
       /// A reference to the chart object
       /// </summary>
       public ExcelChart Chart
       {
           get
           {
               return _chart;
           }
       }
       #region "Add Series"

       /// <summary>
       /// Add a new serie to the chart. Do not apply to pivotcharts.
       /// </summary>
       /// <param name="Serie">The Y-Axis range</param>
       /// <param name="XSerie">The X-Axis range</param>
       /// <returns></returns>
       public virtual ExcelChartSerie Add(ExcelRangeBase Serie, ExcelRangeBase XSerie)
       {
           if (_chart.PivotTableSource != null)
           {
               throw (new InvalidOperationException("Can't add a serie to a pivotchart"));
           }
           return AddSeries(Serie.FullAddressAbsolute, XSerie.FullAddressAbsolute,"");
       }
       /// <summary>
       /// Add a new serie to the chart.Do not apply to pivotcharts.
       /// </summary>
       /// <param name="SerieAddress">The Y-Axis range</param>
       /// <param name="XSerieAddress">The X-Axis range</param>
       /// <returns></returns>
       public virtual ExcelChartSerie Add(string SerieAddress, string XSerieAddress)
       {
           if (_chart.PivotTableSource != null)
           {
               throw (new InvalidOperationException("Can't add a serie to a pivotchart"));
           }
           return AddSeries(SerieAddress, XSerieAddress, "");
       }
       internal protected ExcelChartSerie AddSeries(string SeriesAddress, string XSeriesAddress, string bubbleSizeAddress)
        {
               XmlElement ser = _node.OwnerDocument.CreateElement("ser", ExcelPackage.schemaChart);
               XmlNodeList node = _node.SelectNodes("c:ser", _ns);
               if (node.Count > 0)
               {
                   _node.InsertAfter(ser, node[node.Count-1]);
               }
               else
               {
                   InserAfter(_node, "c:varyColors,c:grouping,c:barDir,c:scatterStyle", ser);
                }
               int idx = FindIndex();
               ser.InnerXml = string.Format("<c:idx val=\"{1}\" /><c:order val=\"{1}\" /><c:tx><c:strRef><c:f></c:f><c:strCache><c:ptCount val=\"1\" /></c:strCache></c:strRef></c:tx>{5}{0}{2}{3}{4}", AddExplosion(Chart.ChartType), idx, AddScatterPoint(Chart.ChartType), AddAxisNodes(Chart.ChartType), AddSmooth(Chart.ChartType), AddMarker(Chart.ChartType));
               ExcelChartSerie serie;
               switch (Chart.ChartType)
               {
                   case eChartType.Bubble:
                   case eChartType.Bubble3DEffect:
                       serie = new ExcelBubbleChartSerie(this, NameSpaceManager, ser, _isPivot)
                       {
                           Bubble3D=Chart.ChartType==eChartType.Bubble3DEffect,
                           Series = SeriesAddress,
                           XSeries = XSeriesAddress,
                           BubbleSize = bubbleSizeAddress                            
                       };
                       break;
                   case eChartType.XYScatter:
                   case eChartType.XYScatterLines:
                   case eChartType.XYScatterLinesNoMarkers:
                   case eChartType.XYScatterSmooth:
                   case eChartType.XYScatterSmoothNoMarkers:
                       serie = new ExcelScatterChartSerie(this, NameSpaceManager, ser, _isPivot);
                       break;
                   case eChartType.Radar:
                   case eChartType.RadarFilled:
                   case eChartType.RadarMarkers:
                       serie = new ExcelRadarChartSerie(this, NameSpaceManager, ser, _isPivot);
                       break;
                   case eChartType.Surface:
                   case eChartType.SurfaceTopView:
                   case eChartType.SurfaceTopViewWireframe:
                   case eChartType.SurfaceWireframe:
                       serie = new ExcelSurfaceChartSerie(this, NameSpaceManager, ser, _isPivot);
                       break;
                   case eChartType.Pie:
                   case eChartType.Pie3D:
                   case eChartType.PieExploded:
                   case eChartType.PieExploded3D:
                   case eChartType.PieOfPie:
                   case eChartType.Doughnut:
                   case eChartType.DoughnutExploded:
                   case eChartType.BarOfPie:
                       serie = new ExcelPieChartSerie(this, NameSpaceManager, ser, _isPivot);
                       break;
                   case eChartType.Line:
                   case eChartType.LineMarkers:
                   case eChartType.LineMarkersStacked:
                   case eChartType.LineMarkersStacked100:
                   case eChartType.LineStacked:
                   case eChartType.LineStacked100:
                       serie = new ExcelLineChartSerie(this, NameSpaceManager, ser, _isPivot);
                       if (Chart.ChartType == eChartType.LineMarkers ||
                           Chart.ChartType == eChartType.LineMarkersStacked ||
                           Chart.ChartType == eChartType.LineMarkersStacked100)
                       {
                           ((ExcelLineChartSerie)serie).Marker = eMarkerStyle.Square;
                       }
                       ((ExcelLineChartSerie)serie).Smooth = ((ExcelLineChart)Chart).Smooth;
                       break;

                   default:
                       serie = new ExcelChartSerie(this, NameSpaceManager, ser, _isPivot);
                       break;
               }               
               serie.Series = SeriesAddress;
               serie.XSeries = XSeriesAddress;                     
           _list.Add(serie);
               return serie;
        }
       bool _isPivot;
       internal void AddPivotSerie(ExcelPivotTable pivotTableSource)
       {
           var r=pivotTableSource.WorkSheet.Cells[pivotTableSource.Address.Address];
           _isPivot = true;
           AddSeries(r.Offset(0, 1, r._toRow - r._fromRow + 1, 1).FullAddressAbsolute, r.Offset(0, 0, r._toRow - r._fromRow + 1, 1).FullAddressAbsolute,"");
       }
       private int FindIndex()
       {    
           int ret = 0, newID=0;
           if (_chart.PlotArea.ChartTypes.Count > 1)
           {
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
                chartType == eChartType.XYScatterSmoothNoMarkers || 
                chartType == eChartType.Bubble ||
                chartType == eChartType.Bubble3DEffect)
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
