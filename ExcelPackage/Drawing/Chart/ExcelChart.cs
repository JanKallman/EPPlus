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
using System.IO;

namespace OfficeOpenXml.Drawing.Chart
{
    #region "Chart Enums"
    public enum eChartType
    {
        Area3D=-4098,
        AreaStacked3D=78,
        AreaStacked1003D=79,
        BarClustered3D= 60,
        BarStacked3D=61,
        BarStacked1003D=62,
        Column3D=-4100,
        ColumnClustered3D=54,
        ColumnStacked3D=55,
        ColumnStacked1003D=56,
        Line3D=-4101,
        Pie3D=-4102,
        PieExploded3D=70,
        Area=1,
        AreaStacked=76,
        AreaStacked100=77,
        BarClustered=57,
        BarOfPie=71,
        BarStacked=58,
        BarStacked100=59,
        Bubble=15,
        Bubble3DEffect=87,
        ColumnClustered=51,
        ColumnStacked=52,
        ColumnStacked100=53,
        ConeBarClustered=102,
        ConeBarStacked=103,
        ConeBarStacked100=104,
        ConeCol=105,
        ConeColClustered=99,
        ConeColStacked=100,
        ConeColStacked100=101,
        CylinderBarClustered=95,
        CylinderBarStacked=96,
        CylinderBarStacked100=97,
        CylinderCol=98,
        CylinderColClustered=92,
        CylinderColStacked=93,
        CylinderColStacked100=94,
        Doughnut=-4120,
        DoughnutExploded=80,
        Line=4,
        LineMarkers=65,
        LineMarkersStacked=66,
        LineMarkersStacked100=67,
        LineStacked=63,
        LineStacked100=64,
        Pie=5,
        PieExploded=69,
        PieOfPie=68,
        PyramidBarClustered=109,
        PyramidBarStacked=110,
        PyramidBarStacked100=111,
        PyramidCol=112,
        PyramidColClustered=106,
        PyramidColStacked=107,
        PyramidColStacked100=108,
        Radar=-4151,
        RadarFilled=82,
        RadarMarkers=81,
        StockHLC=88,
        StockOHLC=89,
        StockVHLC=90,
        StockVOHLC=91,
        Surface=83,
        SurfaceTopView=85,
        SurfaceTopViewWireframe=86,
        SurfaceWireframe=84,
        XYScatter=-4169,
        XYScatterLines=74,
        XYScatterLinesNoMarkers=75,
        XYScatterSmooth=72,
        XYScatterSmoothNoMarkers=73
}
    public enum eDirection
    {
        Column,
        Bar
    }
    public enum eGrouping
    {
        Standard,
        Clustered,
        Stacked,
        PercentStacked
    }
    public enum eShape
    {
        Box,
        Cone,
        ConeToMax,
        Cylinder,
        Pyramid,
        PyramidToMax
    }
    public enum eScatterStyle
    {
        LineMarker,
        SmoothMarker,
     }
    public enum ePieType
    {
        Bar,
        Pie
    }
    public enum eLabelPosition
    {
        BestFit,
        Left,
        Right,
        Center,
        Top,
        Bottom,
        InBase,
        InEnd,
        OutEnd
    }
    public enum eTickLabelPosition
    {
        High,
        Low,
        NextTo,
        None
    }
    /// <summary>
    /// Markerstyle
    /// </summary>
    public enum eMarkerStyle
    {
        Circle,
        Dash,
        Diamond,
        Dot,
        None,
        Picture,
        Plus,
        Square,
        Star,
        Triangle,
        X
    }
    /// <summary>
    /// The build in style of the chart.
    /// </summary>
    public enum eChartStyle
    {
        None,
        Style1,
        Style2,
        Style3,
        Style4,
        Style5,
        Style6,
        Style7,
        Style8,
        Style9,
        Style10,
        Style11,
        Style12,
        Style13,
        Style14,
        Style15,
        Style16,
        Style17,
        Style18,
        Style19,
        Style20,
        Style21,
        Style22,
        Style23,
        Style24,
        Style25,
        Style26,
        Style27,
        Style28,
        Style29,
        Style30,
        Style31,
        Style32,
        Style33,
        Style34,
        Style35,
        Style36,
        Style37,
        Style38,
        Style39,
        Style40,
        Style41,
        Style42,
        Style43,
        Style44,
        Style45,
        Style46,
        Style47,
        Style48
    }
    #endregion
    /// <summary>
   /// Base class for Chart object.
   /// </summary>
    public class ExcelChart : ExcelDrawing
    {
       const string rootPath = "c:chartSpace/c:chart/c:plotArea";
       //string _chartPath;
       ExcelChartSeries _chartSeries;
       internal ExcelChartAxis[] _axis;
       protected XmlHelper _chartXmlHelper;
       #region "Constructors"
       ///// <summary>
       ///// Read the chart from XML
       ///// </summary>
       ///// <param name="drawings">Drawings collection for a worksheet</param>
       ///// <param name="node">Topnode for drawings</param>
       //internal ExcelChart(ExcelDrawings drawings, XmlNode node) :
       //    base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
       // {
       //     XmlNode chartNode = node.SelectSingleNode(string.Format("xdr:graphicFrame/a:graphic/a:graphicData/c:chart"), drawings.NameSpaceManager);
       //     if (chartNode != null)
       //     {
       //         PackageRelationship drawingRelation = drawings.Part.GetRelationship(chartNode.Attributes["r:id"].Value);
       //         UriChart = PackUriHelper.ResolvePartUri(drawings.UriDrawing, drawingRelation.TargetUri);

       //         Part = drawings.Part.Package.GetPart(UriChart);
       //         ChartXml = new XmlDocument();
       //         ChartXml.Load(Part.GetStream());
       //         _chartNode = ChartXml.SelectSingleNode(string.Format("c:chartSpace/c:chart/c:plotArea/{0}",GetChartNodeText()),NameSpaceManager);
       //         InitChartLoad(drawings, _chartNode);
       //     }
       //     else
       //     {
       //         ChartXml = null;
       //     }
       // }
       internal ExcelChart(ExcelDrawings drawings, XmlNode node, eChartType type) :
           base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
       {
           ChartType = type;
           CreateNewChart(drawings, type, null);

           Init(drawings, _chartNode);

           _chartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, _chartNode /*ChartXml.SelectSingleNode(_chartPath, drawings.NameSpaceManager)*/);

           SetTypeProperties(drawings);
           LoadAxis();
       }
       internal ExcelChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart) :
           base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
       {
           ChartType = type;
           CreateNewChart(drawings, type, topChart);

           Init(drawings, _chartNode);

           _chartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, _chartNode);

           SetTypeProperties(drawings);
           if (topChart == null)
               LoadAxis();
           else
           {
               _axis = topChart.Axis;
               if (_axis.Length > 0)
               {
                   XAxis = _axis[0];
                   YAxis = _axis[1];
               }
           }
       }
       internal ExcelChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, PackagePart part, XmlDocument chartXml, XmlNode chartNode) :
           base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
       {
           UriChart = uriChart;
           Part = part;
           ChartXml = chartXml;
           _chartNode = chartNode;
           InitChartLoad(drawings, chartNode);
       }
       internal ExcelChart(ExcelChart topChart, XmlNode chartNode) :
           base(topChart._drawings, topChart.TopNode, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
       {
           UriChart = topChart.UriChart;
           Part = topChart.Part;
           ChartXml = topChart.ChartXml;
           _plotArea = topChart.PlotArea;
           _chartNode = chartNode;

           InitChartLoad(topChart._drawings, chartNode);
       }
       private void InitChartLoad(ExcelDrawings drawings, XmlNode chartNode)
       {
           //SetChartType();
           Init(drawings, chartNode);
           _chartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, _chartNode /*ChartXml.SelectSingleNode(_chartPath, drawings.NameSpaceManager)*/);
           LoadAxis();
       }

       private void Init(ExcelDrawings drawings, XmlNode chartNode)
       {
           _chartXmlHelper = new XmlHelper(drawings.NameSpaceManager, chartNode);
           _chartXmlHelper.SchemaNodeOrder = new string[] { "view3D", "plotArea", "barDir", "grouping", "varyColors", "ser", "dLbls", "shape", "legend", "axId" };
           //if (string.IsNullOrEmpty(chartNodeText))
           //{
           //    chartNodeText = GetChartNodeText();
           //}
          //_chartPath = rootPath + "/" + chartNodeText;
          // _groupingPath = string.Format(_groupingPath, chartNodeText);
          // _varyColorsPath = string.Format(_varyColorsPath, chartNodeText);
           WorkSheet = drawings.Worksheet;
       }
       #endregion
       #region "Private functions"
       private void SetTypeProperties(ExcelDrawings drawings)
       {
           /******* Grouping *******/
           if (IsTypeClustered())
           {
               Grouping = eGrouping.Clustered;
           }
           else if (
               IsTypeStacked())
           {
               Grouping = eGrouping.Stacked;
           }
           else if (
              IsTypePercentStacked())
           {
               Grouping = eGrouping.PercentStacked;
           }

           /***** 3D Perspective *****/
           if (IsType3D())             
           {
               View3D.Perspective = 30;    //Default to 30
               if (IsTypePieDoughnut())
               {
                   View3D.RotX=30;
               }
           }
       }
       private void CreateNewChart(ExcelDrawings drawings, eChartType type, ExcelChart topChart)
       {
           if (topChart == null)
           {
               XmlElement graphFrame = TopNode.OwnerDocument.CreateElement("graphicFrame", ExcelPackage.schemaSheetDrawings);
               graphFrame.SetAttribute("macro", "");
               TopNode.AppendChild(graphFrame);
               graphFrame.InnerXml = "<xdr:nvGraphicFramePr><xdr:cNvPr id=\"2\" name=\"Chart 1\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\" /> <a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\" />   </a:graphicData>  </a:graphic>";
               TopNode.AppendChild(TopNode.OwnerDocument.CreateElement("clientData", ExcelPackage.schemaSheetDrawings));

               Package package = drawings.Worksheet.xlPackage.Package;
               UriChart = GetNewUri(package, "/xl/charts/chart{0}.xml");

               ChartXml = new XmlDocument();
               ChartXml.PreserveWhitespace = ExcelPackage.preserveWhitespace;
               ChartXml.LoadXml(ChartStartXml(type));

               // save it to the package
               Part = package.CreatePart(UriChart, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", _drawings._package.Compression);

               StreamWriter streamChart = new StreamWriter(Part.GetStream(FileMode.Create, FileAccess.Write));
               ChartXml.Save(streamChart);
               streamChart.Close();
               package.Flush();

               PackageRelationship chartRelation = drawings.Part.CreateRelationship(UriChart, TargetMode.Internal, ExcelPackage.schemaRelationships + "/chart");
               graphFrame.SelectSingleNode("a:graphic/a:graphicData/c:chart", NameSpaceManager).Attributes["r:id"].Value = chartRelation.Id;
               package.Flush();
               _chartNode = ChartXml.SelectSingleNode(string.Format("c:chartSpace/c:chart/c:plotArea/{0}", GetChartNodeText()), NameSpaceManager);
           }
           else
           {
               ChartXml = topChart.ChartXml;
               Part = topChart.Part;
               _plotArea = topChart.PlotArea;
               UriChart = topChart.UriChart;
               _axis = topChart._axis;

               XmlNode preNode = _plotArea.ChartTypes[_plotArea.ChartTypes.Count - 1].ChartNode;
               _chartNode = ((XmlDocument)ChartXml).CreateElement(GetChartNodeText(), ExcelPackage.schemaChart);
               preNode.ParentNode.InsertAfter(_chartNode, preNode);
               if (topChart.Axis.Length == 0)
               {
                   AddAxis();
               }
               string serieXML = GetChartSerieStartXml(type, int.Parse(topChart.Axis[0].Id), int.Parse(topChart.Axis[1].Id));
               _chartNode.InnerXml = serieXML;
           }
       }
       private void LoadAxis()
       {
           XmlNodeList nl = _chartNode.SelectNodes("c:axId", NameSpaceManager);
           List<ExcelChartAxis> l = new List<ExcelChartAxis>();
           foreach (XmlNode node in nl)
           {
               string id = node.Attributes["val"].Value;
               var axNode = ChartXml.SelectNodes(rootPath + string.Format("/*/c:axId[@val=\"{0}\"]", id), NameSpaceManager);
               if (axNode != null && axNode.Count>1)
               {
                   foreach (XmlNode axn in axNode)
                   {
                       if (axn.ParentNode.LocalName.EndsWith("Ax"))
                       {
                           XmlNode axisNode = axNode[1].ParentNode;
                           ExcelChartAxis ax = new ExcelChartAxis(NameSpaceManager, axisNode);
                           l.Add(ax);
                       }
                   }
               }
           }
           _axis = l.ToArray();

           if(_axis.Length > 0) XAxis = _axis[0];
           if (_axis.Length > 1) YAxis = _axis[1];
       }
       private void SetChartType()
       {
           ChartType = 0;
           //_plotArea = new ExcelChartPlotArea(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:plotArea", NameSpaceManager));
           int pos=0;
           foreach (XmlElement n in ChartXml.SelectSingleNode(rootPath, _drawings.NameSpaceManager).ChildNodes)
           {
               if (pos == 0)
               {
                   ChartType = GetChartType(n.Name);
                   if (ChartType != 0)
                   {
                       //_chartPath = rootPath + "/" + n.Name;
                       PlotArea.ChartTypes.Add(this);
                   }
               }
               else
               {
                   var chartSerieType = GetChart(_drawings, TopNode/*, n*/);
                   chartSerieType = GetChart(n, _drawings, TopNode, UriChart, Part, ChartXml, null);
                   PlotArea.ChartTypes.Add(chartSerieType);
                   //var chartType = GetChartType(n.Name);
               }
               if (ChartType != 0)
               {
                   pos++;
               }
           }
       }

       private eChartType GetChartType(string name)
       {
           
           switch (name)
           {
               case "c:area3DChart":
                   return eChartType.Area3D;
               case "c:areaChart":
                   return eChartType.Area;
               case "c:barChart":
                   return ChartType = eChartType.BarClustered;
               case "c:bar3DChart":
                   return eChartType.BarClustered3D;
               case "c:bubbleChart":
                   return eChartType.Bubble;
               case "c:doughnutChart":
                   return eChartType.Doughnut;
               case "c:lineChart":
                   return eChartType.Line;
               case "c:line3DChart":
                   return eChartType.Line3D;
               case "c:pie3DChart":
                   return eChartType.Pie3D;
               case "c:pieChart":
                   return eChartType.Pie;
               case "c:radarChart":
                   return eChartType.Radar;
               case "c:scatterChart":
                   return eChartType.XYScatter;
               case "c:surface3DChart":
               case "c:surfaceChart":
                   return eChartType.Surface;
               case "c:stockChart":
                   return eChartType.StockHLC;
               default:
                   return 0;
           }           
       }
       #region "Xml init Functions"
       private string ChartStartXml(eChartType type)
       {
           StringBuilder xml=new StringBuilder();
           int axID=1;
           int xAxID=2;

           xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
           xml.AppendFormat("<c:chartSpace xmlns:c=\"{0}\" xmlns:a=\"{1}\" xmlns:r=\"{2}\">", ExcelPackage.schemaChart, ExcelPackage.schemaDrawings, ExcelPackage.schemaRelationships);       
           xml.Append("<c:chart>");
           xml.AppendFormat("{0}<c:plotArea><c:layout/>",AddPerspectiveXml(type));

           string chartNodeText = GetChartNodeText();
           xml.AppendFormat("<{0}>", chartNodeText);
           xml.Append(GetChartSerieStartXml(type, axID, xAxID));
           xml.AppendFormat("</{0}>", chartNodeText);

           //Axis
           if (!IsTypePieDoughnut())
           {
               xml.AppendFormat("<c:{0}><c:axId val=\"{1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{2}\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:{0}>", AddAxType(), axID, xAxID);
               xml.AppendFormat("<c:valAx><c:axId val=\"{1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"l\"/><c:majorGridlines/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx>", axID, xAxID);
           }
           
           xml.AppendFormat("</c:plotArea><c:legend><c:legendPos val=\"r\"/><c:layout/></c:legend><c:plotVisOnly val=\"1\"/></c:chart>", axID, xAxID);

           xml.Append("<c:printSettings><c:headerFooter/><c:pageMargins b=\"0.75\" l=\"0.7\" r=\"0.7\" t=\"0.75\" header=\"0.3\" footer=\"0.3\"/><c:pageSetup/></c:printSettings></c:chartSpace>");          
           return xml.ToString();
       }

       private string GetChartSerieStartXml(eChartType type, int axID, int xAxID)
       {
           StringBuilder xml = new StringBuilder();

           xml.Append(AddScatterType(type));
           xml.Append(AddVaryColors());
           xml.Append(AddBarDir(type));
           xml.Append(AddGrouping());
           xml.Append(AddHasMarker(type));
           xml.Append(AddShape(type));
           xml.Append(AddFirstSliceAng(type));
           xml.Append(AddHoleSize(type));
           xml.Append(AddAxisId(axID, xAxID));

           return xml.ToString();
       }
       private string AddAxisId(int axID,int xAxID)
       {
           if (!IsTypePieDoughnut())
           {
               return string.Format("<c:axId val=\"{0}\"/><c:axId val=\"{1}\"/>",axID, xAxID);
           }
           else
           {
               return "";
           }
       }
       private string AddAxType()
       {
           switch(ChartType)
           {
               case eChartType.XYScatter:
               case eChartType.XYScatterLines:
               case eChartType.XYScatterLinesNoMarkers:
               case eChartType.XYScatterSmooth:
               case eChartType.XYScatterSmoothNoMarkers:
                   return "valAx";
               default:
                   return "catAx";
           }
       }
       private string AddScatterType(eChartType type)
       {
           if (type == eChartType.XYScatter ||
               type == eChartType.XYScatterLines ||
               type == eChartType.XYScatterLinesNoMarkers ||
               type == eChartType.XYScatterSmooth ||
               type == eChartType.XYScatterSmoothNoMarkers)
           {
               return "<c:scatterStyle val=\"\" />";
           }
           else
           {
               return "";
           }
       }
       private string AddGrouping()
       {
           //IsTypeClustered() || IsTypePercentStacked() || IsTypeStacked() || 
           if(IsTypeShape())
           {
               return "<c:grouping val=\"standard\"/>";
           }
           else
           {
               return "";
           }
       }
       private string AddHoleSize(eChartType type)
       {
           if (type == eChartType.Doughnut ||
               type == eChartType.DoughnutExploded)
           {
               return "<c:holeSize val=\"50\" />";
           }
           else
           {
               return "";
           }
       }
       private string AddFirstSliceAng(eChartType type)
       {
           if (type == eChartType.Doughnut ||
               type == eChartType.DoughnutExploded)
           {
               return "<c:firstSliceAng val=\"0\" />";
           }
           else
           {
               return "";
           }
       }
       private string AddVaryColors()
       {
           if (IsTypePieDoughnut())
           {
               return "<c:varyColors val=\"1\" />";
           }
           else
           {
               return "";
           }
       }
       private string AddHasMarker(eChartType type)
       {
           if (type == eChartType.LineMarkers ||
               type == eChartType.LineMarkersStacked ||
               type == eChartType.LineMarkersStacked100 ||
               type == eChartType.XYScatterLines ||
               type == eChartType.XYScatterSmooth)
           {
               return "<c:marker val=\"1\"/>";
           }
           else
           {
               return "";
           }
       }
       private string AddShape(eChartType type)
       {
           if (IsTypeShape())
           {
               return "<c:shape val=\"box\" />";
           }
           else
           {
               return "";
           }
       }
       private string AddBarDir(eChartType type)
       {
           if (IsTypeShape())
            {
                return "<c:barDir val=\"col\" />";
            }
            else
            {
                return "";
            }
       }
       private string AddPerspectiveXml(eChartType type)
        {
 	        //Add for 3D sharts
            if (IsType3D())
            {
                return "<c:view3D><c:perspective val=\"30\" /></c:view3D>";
            }
            else
            {
                return "";
            }
        }
       #endregion
       #endregion
       #region "Chart type functions
       internal static bool IsType3D(eChartType chartType)
        {
            return chartType == eChartType.Area3D ||
                            chartType == eChartType.AreaStacked3D ||
                            chartType == eChartType.AreaStacked1003D ||
                            chartType == eChartType.BarClustered3D ||
                            chartType == eChartType.BarStacked3D ||
                            chartType == eChartType.BarStacked1003D ||
                            chartType == eChartType.Column3D ||
                            chartType == eChartType.ColumnClustered3D ||
                            chartType == eChartType.ColumnStacked3D ||
                            chartType == eChartType.ColumnStacked1003D ||
                            chartType == eChartType.Line3D ||
                            chartType == eChartType.Pie3D ||
                            chartType == eChartType.PieExploded3D ||
                            chartType == eChartType.Bubble3DEffect ||
                            chartType == eChartType.ConeBarClustered ||
                            chartType == eChartType.ConeBarStacked ||
                            chartType == eChartType.ConeBarStacked100 ||
                            chartType == eChartType.ConeCol ||
                            chartType == eChartType.ConeColClustered ||
                            chartType == eChartType.ConeColStacked ||
                            chartType == eChartType.ConeColStacked100 ||
                            chartType == eChartType.CylinderBarClustered ||
                            chartType == eChartType.CylinderBarStacked ||
                            chartType == eChartType.CylinderBarStacked100 ||
                            chartType == eChartType.CylinderCol ||
                            chartType == eChartType.CylinderColClustered ||
                            chartType == eChartType.CylinderColStacked ||
                            chartType == eChartType.CylinderColStacked100 ||
                            chartType == eChartType.PyramidBarClustered ||
                            chartType == eChartType.PyramidBarStacked ||
                            chartType == eChartType.PyramidBarStacked100 ||
                            chartType == eChartType.PyramidCol ||
                            chartType == eChartType.PyramidColClustered ||
                            chartType == eChartType.PyramidColStacked ||
                            chartType == eChartType.PyramidColStacked100 /*||
                            chartType == eChartType.Doughnut ||
                            chartType == eChartType.DoughnutExploded*/;
        }
       internal protected bool IsType3D()
       {
            return IsType3D(ChartType);
       }
       protected bool IsTypeShape()
       {
            return ChartType == eChartType.BarClustered3D ||
                    ChartType == eChartType.BarStacked3D ||
                    ChartType == eChartType.BarStacked1003D ||
                    ChartType == eChartType.BarClustered3D ||
                    ChartType == eChartType.BarStacked3D ||
                    ChartType == eChartType.BarStacked1003D ||
                    ChartType == eChartType.Column3D ||
                    ChartType == eChartType.ColumnClustered3D ||
                    ChartType == eChartType.ColumnStacked3D ||
                    ChartType == eChartType.ColumnStacked1003D ||
                //ChartType == eChartType.3DPie ||
                //ChartType == eChartType.3DPieExploded ||
                    ChartType == eChartType.Bubble3DEffect ||
                    ChartType == eChartType.ConeBarClustered ||
                    ChartType == eChartType.ConeBarStacked ||
                    ChartType == eChartType.ConeBarStacked100 ||
                    ChartType == eChartType.ConeCol ||
                    ChartType == eChartType.ConeColClustered ||
                    ChartType == eChartType.ConeColStacked ||
                    ChartType == eChartType.ConeColStacked100 ||
                    ChartType == eChartType.CylinderBarClustered ||
                    ChartType == eChartType.CylinderBarStacked ||
                    ChartType == eChartType.CylinderBarStacked100 ||
                    ChartType == eChartType.CylinderCol ||
                    ChartType == eChartType.CylinderColClustered ||
                    ChartType == eChartType.CylinderColStacked ||
                    ChartType == eChartType.CylinderColStacked100 ||
                    ChartType == eChartType.PyramidBarClustered ||
                    ChartType == eChartType.PyramidBarStacked ||
                    ChartType == eChartType.PyramidBarStacked100 ||
                    ChartType == eChartType.PyramidCol ||
                    ChartType == eChartType.PyramidColClustered ||
                    ChartType == eChartType.PyramidColStacked ||
                    ChartType == eChartType.PyramidColStacked100; //||
                    //ChartType == eChartType.Doughnut ||
                    //ChartType == eChartType.DoughnutExploded;
        }
        protected bool IsTypePercentStacked()
        {
            return ChartType == eChartType.AreaStacked100 ||
                           ChartType == eChartType.BarStacked100 ||
                           ChartType == eChartType.ConeBarStacked100 ||
                           ChartType == eChartType.ConeColStacked100 ||
                           ChartType == eChartType.CylinderBarStacked100 ||
                           ChartType == eChartType.CylinderColStacked ||
                           ChartType == eChartType.LineMarkersStacked100 ||
                           ChartType == eChartType.LineStacked100 ||
                           ChartType == eChartType.PyramidBarStacked100 ||
                           ChartType == eChartType.PyramidColStacked100;
        }
        protected bool IsTypeStacked()
        {
            return ChartType == eChartType.AreaStacked ||
                           ChartType == eChartType.BarStacked ||
                           ChartType == eChartType.ColumnStacked3D ||
                           ChartType == eChartType.ConeBarStacked ||
                           ChartType == eChartType.ConeColStacked ||
                           ChartType == eChartType.CylinderBarStacked ||
                           ChartType == eChartType.CylinderColStacked ||
                           ChartType == eChartType.LineMarkersStacked ||
                           ChartType == eChartType.LineStacked ||
                           ChartType == eChartType.PyramidBarStacked ||
                           ChartType == eChartType.PyramidColStacked;
        }
        protected bool IsTypeClustered()
        {
            return ChartType == eChartType.BarClustered ||
                           ChartType == eChartType.BarClustered3D ||
                           ChartType == eChartType.ColumnClustered3D ||
                           ChartType == eChartType.ColumnClustered ||
                           ChartType == eChartType.ConeBarClustered ||
                           ChartType == eChartType.ConeColClustered ||
                           ChartType == eChartType.CylinderBarClustered ||
                           ChartType == eChartType.CylinderColClustered ||
                           ChartType == eChartType.PyramidBarClustered ||
                           ChartType == eChartType.PyramidColClustered;
        }
        protected bool IsTypePieDoughnut()
        {
            return ChartType == eChartType.Pie ||
                           ChartType == eChartType.PieExploded ||
                           ChartType == eChartType.PieOfPie ||
                           ChartType == eChartType.Pie3D ||
                           ChartType == eChartType.PieExploded3D ||
                           ChartType == eChartType.BarOfPie ||
                           ChartType == eChartType.Doughnut ||
                           ChartType == eChartType.DoughnutExploded;
        }
        #endregion
       /// <summary>
       /// Get the name of the chart node
       /// </summary>
       /// <returns>The name</returns>
        protected string GetChartNodeText()
        {
            switch (ChartType)
            {
                case eChartType.Area3D:
                case eChartType.AreaStacked3D:
                case eChartType.AreaStacked1003D:
                    return "c:area3DChart";
                case eChartType.Area:
                case eChartType.AreaStacked:
                case eChartType.AreaStacked100:
                    return "c:areaChart";
                case eChartType.BarClustered:
                case eChartType.BarStacked:
                case eChartType.BarStacked100:
                case eChartType.ColumnClustered:
                case eChartType.ColumnStacked:
                case eChartType.ColumnStacked100:
                    return "c:barChart";
                case eChartType.BarClustered3D:
                case eChartType.BarStacked3D:
                case eChartType.BarStacked1003D:
                case eChartType.ColumnClustered3D:
                case eChartType.ColumnStacked3D:
                case eChartType.ColumnStacked1003D:
                case eChartType.ConeBarClustered:
                case eChartType.ConeBarStacked:
                case eChartType.ConeBarStacked100:
                case eChartType.ConeCol:
                case eChartType.ConeColClustered:
                case eChartType.ConeColStacked:
                case eChartType.ConeColStacked100:
                case eChartType.CylinderBarClustered:
                case eChartType.CylinderBarStacked:
                case eChartType.CylinderBarStacked100:
                case eChartType.CylinderCol:
                case eChartType.CylinderColClustered:
                case eChartType.CylinderColStacked:
                case eChartType.CylinderColStacked100:
                case eChartType.PyramidBarClustered:
                case eChartType.PyramidBarStacked:
                case eChartType.PyramidBarStacked100:
                case eChartType.PyramidCol:
                case eChartType.PyramidColClustered:
                case eChartType.PyramidColStacked:
                case eChartType.PyramidColStacked100:
                    return "c:bar3DChart";
                case eChartType.Bubble:
                    return "c:bubbleChart";
                case eChartType.Doughnut:
                case eChartType.DoughnutExploded:
                    return "c:doughnutChart";
                case eChartType.Line:
                case eChartType.LineMarkers:
                case eChartType.LineMarkersStacked:
                case eChartType.LineMarkersStacked100:
                case eChartType.LineStacked:
                case eChartType.LineStacked100:
                    return "c:lineChart";
                case eChartType.Line3D:
                    return "c:line3DChart";
                case eChartType.Pie:
                case eChartType.PieExploded:
                    return "c:pieChart";
                case eChartType.BarOfPie:
                case eChartType.PieOfPie:
                    return "c:ofPieChart";
                case eChartType.Pie3D:
                case eChartType.PieExploded3D:
                    return "c:pie3DChart";
                case eChartType.Radar:
                case eChartType.RadarFilled:
                case eChartType.RadarMarkers:
                    return "c:radarChart";
                case eChartType.XYScatter:
                case eChartType.XYScatterLines:
                case eChartType.XYScatterLinesNoMarkers:
                case eChartType.XYScatterSmooth:
                case eChartType.XYScatterSmoothNoMarkers:
                    return "c:scatterChart";
                case eChartType.Surface:
                    return "c:surfaceChart";
                case eChartType.StockHLC:
                    return "c:stockChart";
                default:
                    throw(new NotImplementedException("Chart type not implemented"));
            }
        }
        /// <summary>
        /// Add a secondary axis
        /// </summary>
        internal void AddAxis()
        {
            XmlElement catAx = ChartXml.CreateElement(string.Format("c:{0}",AddAxType()), ExcelPackage.schemaChart);
            int axID;
            if (_axis.Length == 0)
            {
                _plotArea.TopNode.AppendChild(catAx);
                axID = 1;
            }
            else
            {
                _axis[0].TopNode.ParentNode.InsertAfter(catAx, _axis[_axis.Length-1].TopNode);
                axID = int.Parse(_axis[0].Id) < int.Parse(_axis[1].Id) ? int.Parse(_axis[1].Id) + 1 : int.Parse(_axis[0].Id) + 1;
            }


            XmlElement valAx = ChartXml.CreateElement("c:valAx", ExcelPackage.schemaChart);
            catAx.ParentNode.InsertAfter(valAx, catAx);

            if (_axis.Length == 0)
            {
                catAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/>", axID, axID+1);
                valAx.InnerXml = string.Format("<c:axId val=\"{1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"l\"/><c:majorGridlines/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{0}\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/>", axID, axID+1);
            }
            else
            {
                catAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"1\" /><c:axPos val=\"b\"/><c:tickLblPos val=\"none\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/>", axID, axID + 1);
                valAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"r\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"max\"/><c:crossBetween val=\"between\"/>", axID + 1, axID);
            }

            if (_axis.Length == 0)
            {
                _axis = new ExcelChartAxis[2];
            }
            else
            {
                ExcelChartAxis[] newAxis = new ExcelChartAxis[_axis.Length + 2];
                Array.Copy(_axis, newAxis, _axis.Length);
                _axis = newAxis;
            }

            _axis[_axis.Length - 2] = new ExcelChartAxis(NameSpaceManager, catAx);
            _axis[_axis.Length - 1] = new ExcelChartAxis(NameSpaceManager, valAx);
            foreach (var chart in _plotArea.ChartTypes)
            {
                chart._axis = _axis;
            }
        }
        internal void RemoveSecondaryAxis()
        {
            throw (new NotImplementedException("Not yet implemented"));
        }
        #region "Properties"
        public ExcelWorksheet WorkSheet { get; internal set; }
        public XmlDocument ChartXml { get; set; }
        public eChartType ChartType { get; set; }
        XmlNode _chartNode = null;
        internal XmlNode ChartNode
        {
            get
            {
                return _chartNode;
            }
        }
        /// <summary>
        /// Titel of the chart
        /// </summary>
        public ExcelChartTitle Title
        {
            get
            {
                if (_title == null)
                {
                    _title = new ExcelChartTitle(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart", NameSpaceManager));
                }
                return _title;
            }
        }
        /// <summary>
        /// Chart series
        /// </summary>
        public ExcelChartSeries Series
        {
            get
            {
                return _chartSeries;
            }
        }
        /// <summary>
        /// An array containg all axis of all Charttypes
        /// </summary>
        public ExcelChartAxis[] Axis
        {
            get
            {
                return _axis;
            }
        }
        /// <summary>
        /// The XAxis
        /// </summary>
        public ExcelChartAxis XAxis
        {
            get;
            private set;
        }
        /// <summary>
        /// The YAxis
        /// </summary>
        public ExcelChartAxis YAxis
        {
            get;
            private set;
        }
        bool _secondaryAxis=false;
        /// <summary>
        /// If true the axis will use the secondary axis.
        /// The chart must contain a least one other chart serie that uses the primary axis.
        /// </summary>
        public bool UseSecondaryAxis
        {
            get
            {
                return _secondaryAxis;
            }
            set
            {
                 if (_secondaryAxis != value)
                {
                    if (value)
                    {
                        if (IsTypePieDoughnut())
                        {
                            throw (new Exception("Pie charts do not support axis"));
                        }
                        else if (HasPrimaryAxis() == false)
                        {
                            throw (new Exception("Can's set to secondary Y axis when no serie uses the primary axis"));
                        }
                        if (Axis.Length == 2)
                        {
                            AddAxis();
                        }
                        var nl = ChartNode.SelectNodes("c:axId", NameSpaceManager);
                        nl[0].Attributes["val"].Value = Axis[2].Id;
                        nl[1].Attributes["val"].Value = Axis[3].Id;
                        XAxis = Axis[2];
                        YAxis = Axis[3];
                    }
                    else
                    {
                        var nl = ChartNode.SelectNodes("c:axId", NameSpaceManager);
                        nl[0].Attributes["val"].Value = Axis[0].Id;
                        nl[1].Attributes["val"].Value = Axis[1].Id;
                        XAxis = Axis[0];
                        YAxis = Axis[1];
                    }
                    _secondaryAxis = value;
                }
            }
        }
        public eChartStyle Style
        {
            get
            {
                XmlNode node = ChartXml.SelectSingleNode("c:chartSpace/c:style/@val", NameSpaceManager);
                if (node == null)
                {
                    return eChartStyle.None;
                }
                else
                {
                    int v;
                    if (int.TryParse(node.Value, out v))
                    {
                        return (eChartStyle)v;
                    }
                    else
                    {
                        return eChartStyle.None;
                    }
                }

            }
            set
            {
                if (value == eChartStyle.None)
                {
                    XmlElement element = ChartXml.SelectSingleNode("c:chartSpace/c:style", NameSpaceManager) as XmlElement;
                    if (element != null)
                    {
                        element.ParentNode.RemoveChild(element);
                    }
                }
                else
                {
                    XmlElement element = ChartXml.CreateElement("c:style", ExcelPackage.schemaChart);
                    element.SetAttribute("val", ((int)value).ToString());
                    XmlElement parent = ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager) as XmlElement;
                    parent.InsertBefore(element, parent.SelectSingleNode("c:chart", NameSpaceManager));
                }
            }
        }
        private bool HasPrimaryAxis()
        {
            if (_plotArea.ChartTypes.Count == 1)
            {
                return false;
            }
            foreach (var chart in _plotArea.ChartTypes)
            {
                if (chart != this)
                {
                    if (chart.UseSecondaryAxis == false && chart.IsTypePieDoughnut()==false)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        ///// <summary>
        ///// Sets position of the axis of a chart-serie
        ///// </summary>
        ///// <param name="XAxis">Left or Right</param>
        ///// <param name="YAxis">Top or Bottom</param>
        //internal void SetAxis(eXAxisPosition XAxis, eYAxisPosition YAxis)
        //{
        //    bool xAxisExists = false, yAxisExists = false;
        //    foreach (var axis in _axis)
        //    {
        //        if (axis.AxisPosition == (eAxisPosition)XAxis)
        //        {
        //            //Found
        //            xAxisExists=true;
        //            if (axis != this.XAxis)
        //            {
        //                CheckRemoveAxis(this.XAxis);
        //                this.XAxis = axis;
        //            }
        //        }
        //        else if (axis.AxisPosition == (eAxisPosition)YAxis)
        //        {
        //            yAxisExists = true;
        //            if (axis != this.YAxis)
        //            {
        //                CheckRemoveAxis(this.YAxis);
        //                this.YAxis = axis;
        //            }
        //        }
        //    }

        //    if (!xAxisExists)
        //    {
        //        if (ExistsAxis(this.XAxis))
        //        {
        //            AddAxis((eAxisPosition)XAxis);
        //            this.XAxis = Axis[Axis.Length - 1];
        //        }
        //        else
        //        {
        //            this.XAxis.AxisPosition = (eAxisPosition)XAxis;
        //        }
        //    }
        //    if (!yAxisExists)
        //    {
        //        if (ExistsAxis(this.XAxis))
        //        {
        //            AddAxis((eAxisPosition)YAxis);
        //            this.YAxis = Axis[Axis.Length - 1];
        //        }
        //        else
        //        {
        //            this.YAxis.AxisPosition = (eAxisPosition)YAxis;
        //        }
        //    }
        //}

        /// <summary>
        /// Remove all axis that are not used any more
        /// </summary>
        /// <param name="excelChartAxis"></param>
        private void CheckRemoveAxis(ExcelChartAxis excelChartAxis)
        {
            if (ExistsAxis(excelChartAxis))
            {
                //Remove the axis
                ExcelChartAxis[] newAxis = new ExcelChartAxis[Axis.Length - 1];
                int pos = 0;
                foreach (var ax in Axis)
                {
                    if (ax != excelChartAxis)
                    {
                        newAxis[pos] = ax;
                    }
                }

                //Update all charttypes.
                foreach (ExcelChart chartType in _plotArea.ChartTypes)
                {
                    chartType._axis = newAxis;
                }
            }
        }

        private bool ExistsAxis(ExcelChartAxis excelChartAxis)
        {
            foreach (ExcelChart chartType in _plotArea.ChartTypes)
            {
                if (chartType != this)
                {
                    if (chartType.XAxis.AxisPosition == excelChartAxis.AxisPosition ||
                       chartType.YAxis.AxisPosition == excelChartAxis.AxisPosition)
                    {
                        //The axis exists
                        return true;
                    }
                }
            }
            return false;
        }
        ExcelChartPlotArea _plotArea = null;
        /// <summary>
        /// Plotarea
        /// </summary>
        public ExcelChartPlotArea PlotArea
        {
            get
            {
                if (_plotArea == null)
                {
                    _plotArea = new ExcelChartPlotArea(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:plotArea", NameSpaceManager), this); 
                }
                return _plotArea;
            }
        }
        ExcelChartLegend _legend = null;
        /// <summary>
        /// Legend
        /// </summary>
        public ExcelChartLegend Legend
        {
            get
            {
                if (_legend == null)
                {
                    _legend = new ExcelChartLegend(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", NameSpaceManager), this);
                }
                return _legend;
            }

        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Border
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace",NameSpaceManager), "c:spPr/a:ln"); 
                }
                return _border;
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Fill
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:spPr");
                }
                return _fill;
            }
        }
        public ExcelView3D View3D
        {
            get
            {
                if (IsType3D())
                {
                    return new ExcelView3D(NameSpaceManager, ChartXml.SelectSingleNode("//c:view3D", NameSpaceManager));
                }
                else
                {
                    throw (new Exception("Charttype does not support 3D"));
                }

            }
        }
        //string _groupingPath = "c:chartSpace/c:chart/c:plotArea/{0}/c:grouping/@val";
        string _groupingPath = "c:grouping/@val";
        public eGrouping Grouping
        {
            get
            {
                return GetGroupingEnum(_chartXmlHelper.GetXmlNodeString(_groupingPath));
            }
            internal set
            {
                _chartXmlHelper.SetXmlNodeString(_groupingPath, GetGroupingText(value));
            }
        }
        //string _varyColorsPath = "c:chartSpace/c:chart/c:plotArea/{0}/c:varyColors/@val";
        string _varyColorsPath = "c:varyColors/@val";
        /// <summary>
        /// If the chart has only one serie this varies the colors for each point.
        /// </summary>
        public bool VaryColors
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeBool(_varyColorsPath);
            }
            set
            {
                if (value)
                {
                    _chartXmlHelper.SetXmlNodeString(_varyColorsPath, "1");
                }
                else
                {
                    _chartXmlHelper.SetXmlNodeString(_varyColorsPath, "0");
                }
            }
        }
        internal PackagePart Part { get; set; }
        internal Uri UriChart { get; set; }
        internal string Id
        {
            get { return ""; }
        }
        ExcelChartTitle _title = null;
        #endregion
       #region "Grouping Enum Translation"
        private string GetGroupingText(eGrouping grouping)
       {
           switch (grouping)
           {
               case eGrouping.Clustered:
                   return "clustered";
               case eGrouping.Stacked:
                   return "stacked";
               case eGrouping.PercentStacked:
                   return "percentStacked";
               default:
                   return "standard";

           }
       }
       private eGrouping GetGroupingEnum(string grouping)
       {
           switch (grouping)
           {
               case "stacked":
                   return eGrouping.Stacked;
               case "percentStacked":
                   return eGrouping.PercentStacked;
               default: //"clustered":               
                   return eGrouping.Clustered;
           }         
       }
       #endregion
       internal static ExcelChart GetChart(ExcelDrawings drawings, XmlNode node/*, XmlNode chartTypeNode*/)
       {
           XmlNode chartNode = node.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart", drawings.NameSpaceManager);
           if (chartNode != null)
           {
               PackageRelationship drawingRelation = drawings.Part.GetRelationship(chartNode.Attributes["r:id"].Value);
               var uriChart = PackUriHelper.ResolvePartUri(drawings.UriDrawing, drawingRelation.TargetUri);

               var part = drawings.Part.Package.GetPart(uriChart);
               var chartXml = new XmlDocument();
               chartXml.Load(part.GetStream());

               ExcelChart topChart = null;
               foreach (XmlElement n in chartXml.SelectSingleNode(rootPath, drawings.NameSpaceManager).ChildNodes)
                {
                    if (topChart == null)
                    {
                        topChart = GetChart(n, drawings, node, uriChart, part, chartXml, null);
                        if(topChart!=null)
                        {
                            topChart.PlotArea.ChartTypes.Add(topChart);
                        }
                    }
                    else
                    {
                        var subChart = GetChart(n, null, null, null, null, null, topChart);
                        if (subChart != null)
                        {
                            topChart.PlotArea.ChartTypes.Add(subChart);
                        }
                    }
                }               
                return topChart;
           }
           else
           {
               return null;
           }           
       }
       internal static ExcelChart GetChart(XmlElement chartNode, ExcelDrawings drawings, XmlNode node,  Uri uriChart, PackagePart part, XmlDocument chartXml, ExcelChart topChart)
       {
           switch (chartNode.LocalName)
           {
               case "area3DChart":
               case "areaChart":
               case "bubbleChart":
               case "radarChart":
               case "surface3DChart":
               case "surfaceChart":
               case "stockChart":
                   if (topChart == null)
                   {
                       return new ExcelChart(drawings, node, uriChart, part, chartXml, chartNode);
                   }
                   else
                   {
                       return new ExcelChart(topChart, chartNode);
                   }
               case "barChart":
               case "bar3DChart":
                   if (topChart == null)
                   {
                       return new ExcelBarChart(drawings, node, uriChart, part, chartXml, chartNode);
                   }
                   else
                   {
                       return new ExcelBarChart(topChart, chartNode);
                   }
               case "doughnutChart":
                   if (topChart == null)
                   {
                       return new ExcelDoughnutChart(drawings, node, uriChart, part, chartXml, chartNode);
                   }
                   else
                   {
                       return new ExcelDoughnutChart(topChart, chartNode);
                   }
               case "pie3DChart":
               case "pieChart":
                   if (topChart == null)
                   {
                       return new ExcelPieChart(drawings, node, uriChart, part, chartXml, chartNode);
                   }
                   else
                   {
                       return new ExcelPieChart(topChart, chartNode);
                   }
           case "ofPieChart":
                   if (topChart == null)
                   {
                       return new ExcelOfPieChart(drawings, node, uriChart, part, chartXml, chartNode);
                   }
                   else
                   {
                       return new ExcelBarChart(topChart, chartNode);
                   }
               case "lineChart":
               case "line3DChart":
                   if (topChart == null)
                   {
                       return new ExcelLineChart(drawings, node, uriChart, part, chartXml, chartNode);
                   }
                   else
                   {
                       return new ExcelLineChart(topChart, chartNode);
                   }
               case "scatterChart":
                   if (topChart == null)
                   {
                       return new ExcelScatterChart(drawings, node, uriChart, part, chartXml, chartNode);
                   }
                   else
                   {
                       return new ExcelScatterChart(topChart, chartNode);
                   }
               default:
                   return null;
           }       
       }
       internal static ExcelChart GetNewChart(ExcelDrawings drawings, XmlNode drawNode, eChartType chartType, ExcelChart topChart)
       {
            switch(chartType)
            {
                case eChartType.Pie:
                case eChartType.PieExploded:
                case eChartType.Pie3D:
                case eChartType.PieExploded3D:
                    return new ExcelPieChart(drawings, drawNode, chartType, topChart);
                case eChartType.BarOfPie:
                case eChartType.PieOfPie:
                    return new ExcelOfPieChart(drawings, drawNode, chartType, topChart);
                case eChartType.Doughnut:
                case eChartType.DoughnutExploded:
                    return new ExcelDoughnutChart(drawings, drawNode, chartType, topChart);
                case eChartType.BarClustered:
                case eChartType.BarStacked:
                case eChartType.BarStacked100:
                case eChartType.BarClustered3D:
                case eChartType.BarStacked3D:
                case eChartType.BarStacked1003D:
                case eChartType.ConeBarClustered:
                case eChartType.ConeBarStacked:
                case eChartType.ConeBarStacked100:
                case eChartType.CylinderBarClustered:
                case eChartType.CylinderBarStacked:
                case eChartType.CylinderBarStacked100:
                case eChartType.PyramidBarClustered:
                case eChartType.PyramidBarStacked:
                case eChartType.PyramidBarStacked100:
                case eChartType.ColumnClustered:
                case eChartType.ColumnStacked:
                case eChartType.ColumnStacked100:
                case eChartType.Column3D:
                case eChartType.ColumnClustered3D:
                case eChartType.ColumnStacked3D:
                case eChartType.ColumnStacked1003D:
                case eChartType.ConeCol:
                case eChartType.ConeColClustered:
                case eChartType.ConeColStacked:
                case eChartType.ConeColStacked100:
                case eChartType.CylinderCol:
                case eChartType.CylinderColClustered:
                case eChartType.CylinderColStacked:
                case eChartType.CylinderColStacked100:
                case eChartType.PyramidCol:
                case eChartType.PyramidColClustered:
                case eChartType.PyramidColStacked:
                case eChartType.PyramidColStacked100:
                    return new ExcelBarChart(drawings, drawNode, chartType, topChart);
                case eChartType.XYScatter:
                case eChartType.XYScatterLines:
                case eChartType.XYScatterLinesNoMarkers:
                case eChartType.XYScatterSmooth:
                case eChartType.XYScatterSmoothNoMarkers:
                    return new ExcelScatterChart(drawings, drawNode, chartType, topChart);
                case eChartType.Line:
                case eChartType.Line3D:
                case eChartType.LineMarkers:
                case eChartType.LineMarkersStacked:
                case eChartType.LineMarkersStacked100:
                case eChartType.LineStacked:
                case eChartType.LineStacked100:
                    return new ExcelLineChart(drawings, drawNode, chartType, topChart);
                default:
                    return new ExcelChart(drawings, drawNode, chartType, topChart);
            }
        }
    }
}
