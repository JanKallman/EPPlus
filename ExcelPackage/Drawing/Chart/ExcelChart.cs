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
    public enum eTickLablePosition
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
    #endregion
    /// <summary>
   /// Provide access to Chart objects.
   /// </summary>
    public class ExcelChart : ExcelDrawing
    {
       const string rootPath = "c:chartSpace/c:chart/c:plotArea";
       string _chartPath;
       ExcelChartSeries _chartSeries;
       ExcelChartAxis[] _axis;
       protected XmlHelper _chartXmlHelper;
       #region "Constructors"
       /// <summary>
       /// Read the chart from XML
       /// </summary>
       /// <param name="drawings">Drawings collection for a worksheet</param>
       /// <param name="node">Topnode for drawings</param>        
        internal ExcelChart(ExcelDrawings drawings, XmlNode node) :
           base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
        {
            XmlNode chartNode = node.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart", drawings.NameSpaceManager);
            if (chartNode != null)
            {
                PackageRelationship drawingRelation = drawings.Part.GetRelationship(chartNode.Attributes["r:id"].Value);
                UriChart = PackUriHelper.ResolvePartUri(drawings.UriDrawing, drawingRelation.TargetUri);

                Part = drawings.Part.Package.GetPart(UriChart);
                ChartXml = new XmlDocument();
                ChartXml.Load(Part.GetStream());
                WorkSheet = drawings.Worksheet;
                SetChartType();
                _chartXmlHelper = new XmlHelper(drawings.NameSpaceManager, ChartXml);
                _chartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, ChartXml.SelectSingleNode(_chartPath, drawings.NameSpaceManager));
                LoadAxis();
            }
            else
            {
                ChartXml = null;
            }
        }
       internal ExcelChart(ExcelDrawings drawings, XmlNode node, eChartType type) :
           base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
       {
           SchemaNodeOrder = new string[] {"view3D", "plotArea", "barDir", "grouping", "ser", "dLbls", "shape", "legend" };
           ChartType = type;
           CreateNewChart(drawings, type);
           _chartPath = rootPath + "/" + GetChartNodeText();
           WorkSheet = drawings.Worksheet;

           string chartNodeText=GetChartNodeText();
           _groupingPath = string.Format(_groupingPath, chartNodeText);

           _chartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, ChartXml.SelectSingleNode(_chartPath, drawings.NameSpaceManager));
           _chartXmlHelper = new XmlHelper(drawings.NameSpaceManager, ChartXml);

           SetTypeProperties(drawings);
           LoadAxis();
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
       private void CreateNewChart(ExcelDrawings drawings, eChartType type)
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
       }
       private void LoadAxis()
       {
           XmlNodeList nl = ChartXml.SelectNodes(_chartPath + "/c:axId", NameSpaceManager);
           List<ExcelChartAxis> l = new List<ExcelChartAxis>();
           foreach (XmlNode node in nl)
           {
               string id = node.Attributes["val"].Value;
               XmlNode axisNode = ChartXml.SelectNodes(rootPath + string.Format("/*/c:axId[@val=\"{0}\"]", id), NameSpaceManager)[1].ParentNode;
               ExcelChartAxis ax = new ExcelChartAxis(NameSpaceManager, axisNode);
               l.Add(ax);
           }
           _axis = l.ToArray();
       }
       private void SetChartType()
       {
           ChartType = 0;
           foreach (XmlNode n in ChartXml.SelectSingleNode(rootPath, _drawings.NameSpaceManager).ChildNodes)
           {
               switch (n.Name)
               {
                   case "c:area3DChart":
                       ChartType = eChartType.Area3D;
                       break;
                   case "c:areaChart":
                       ChartType = eChartType.Area;
                       break;
                   case "c:barChart":
                       ChartType = eChartType.BarClustered;
                       break;
                   case "c:bar3DChart":
                       ChartType = eChartType.BarClustered3D;
                       break;
                   case "c:bubbleChart":
                       ChartType = eChartType.Bubble;
                       break;
                   case "c:doughnutChart":
                       ChartType = eChartType.Doughnut;
                       break;
                   case "c:lineChart":
                       ChartType = eChartType.Line;
                       break;
                   case "c:line3DChart":
                       ChartType = eChartType.Line3D;
                       break;
                   case "c:pie3DChart":
                       ChartType = eChartType.Pie3D;
                       break;
                   case "c:pieChart":
                       ChartType = eChartType.Pie;
                       break;
                   case "c:radarChart":
                       ChartType = eChartType.Radar;
                       break;
                   case "c:scatterChart":
                       ChartType = eChartType.XYScatter;
                       break;
                   case "c:surface3DChart":
                   case "c:surfaceChart":
                       ChartType = eChartType.Surface;
                       break;
                   case "c:stockChart":
                       ChartType = eChartType.StockHLC;
                       break;
               }
               if (ChartType != 0)
               {
                   _chartPath = rootPath + "/" + n.Name;
                   return;
               }
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

           xml.Append(AddScatterType(type));
           xml.Append(AddVaryColors());
           xml.Append(AddBarDir(type));
           xml.Append(AddGrouping());
           xml.Append(AddHasMarker(type));
           xml.Append(AddShape(type));
           xml.Append(AddFirstSliceAng(type));
           xml.Append(AddHoleSize(type));
           xml.Append(AddAxisId(axID, xAxID));

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
       protected bool IsType3D()
        {
            return ChartType == eChartType.Area3D ||
                            ChartType == eChartType.AreaStacked3D ||
                            ChartType == eChartType.AreaStacked1003D ||
                            ChartType == eChartType.BarClustered3D ||
                            ChartType == eChartType.BarStacked3D ||
                            ChartType == eChartType.BarStacked1003D ||
                            ChartType == eChartType.Column3D ||
                            ChartType == eChartType.ColumnClustered3D ||
                            ChartType == eChartType.ColumnStacked3D ||
                            ChartType == eChartType.ColumnStacked1003D ||
                            ChartType == eChartType.Line3D ||
                            ChartType == eChartType.Pie3D ||
                            ChartType == eChartType.PieExploded3D ||
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
                            ChartType == eChartType.PyramidColStacked100 ||
                            ChartType == eChartType.Doughnut ||
                            ChartType == eChartType.DoughnutExploded;
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
       #region "Properties"
        public ExcelWorksheet WorkSheet { get; internal set; }
        public XmlDocument ChartXml { get; set; }
        public eChartType ChartType { get; set; }
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
        /// Axis 
        /// </summary>
        public ExcelChartAxis[] Axis
        {
            get
            {
                return _axis;
            }
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
                    _plotArea = new ExcelChartPlotArea(NameSpaceManager, _chartXmlHelper.TopNode.SelectSingleNode("c:chartSpace/c:chart/c:plotArea", NameSpaceManager)); 
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
                    _legend = new ExcelChartLegend(NameSpaceManager, _chartXmlHelper.TopNode.SelectSingleNode("c:chartSpace/c:chart/c:legend", NameSpaceManager), this);
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
                    _border = new ExcelDrawingBorder(NameSpaceManager, _chartXmlHelper.TopNode.SelectSingleNode("c:chartSpace",NameSpaceManager), "c:spPr/a:ln"); 
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
                    _fill = new ExcelDrawingFill(NameSpaceManager, _chartXmlHelper.TopNode.SelectSingleNode("c:chartSpace", NameSpaceManager), "c:spPr");
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
        string _groupingPath = "c:chartSpace/c:chart/c:plotArea/{0}/c:grouping/@val";
        public eGrouping Grouping
        {
            get
            {
                return GetGroupingEnum(_chartXmlHelper.GetXmlNode(_groupingPath));
            }
            internal set
            {
                _chartXmlHelper.SetXmlNode(_groupingPath, GetGroupingText(value));
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
    }
}
