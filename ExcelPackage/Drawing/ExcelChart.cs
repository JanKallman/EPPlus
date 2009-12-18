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
using System.IO;

namespace OfficeOpenXml.Drawing
{
    #region "Chart Enums"
    public enum eChartType
    {
        xl3DArea=-4098,
        xl3DAreaStacked=78,
        xl3DAreaStacked100=79,
        xl3DBarClustered=60,
        xl3DBarStacked=61,
        xl3DBarStacked100=62,
        xl3DColumn=-4100,
        xl3DColumnClustered=54,
        xl3DColumnStacked=55,
        xl3DColumnStacked100=56,
        xl3DLine=-4101,
        xl3DPie=-4102,
        xl3DPieExploded=70,
        xlArea=1,
        xlAreaStacked=76,
        xlAreaStacked100=77,
        xlBarClustered=57,
        xlBarOfPie=71,
        xlBarStacked=58,
        xlBarStacked100=59,
        xlBubble=15,
        xlBubble3DEffect=87,
        xlColumnClustered=51,
        xlColumnStacked=52,
        xlColumnStacked100=53,
        xlConeBarClustered=102,
        xlConeBarStacked=103,
        xlConeBarStacked100=104,
        xlConeCol=105,
        xlConeColClustered=99,
        xlConeColStacked=100,
        xlConeColStacked100=101,
        xlCylinderBarClustered=95,
        xlCylinderBarStacked=96,
        xlCylinderBarStacked100=97,
        xlCylinderCol=98,
        xlCylinderColClustered=92,
        xlCylinderColStacked=93,
        xlCylinderColStacked100=94,
        xlDoughnut=-4120,
        xlDoughnutExploded=80,
        xlLine=4,
        xlLineMarkers=65,
        xlLineMarkersStacked=66,
        xlLineMarkersStacked100=67,
        xlLineStacked=63,
        xlLineStacked100=64,
        xlPie=5,
        xlPieExploded=69,
        xlPieOfPie=68,
        xlPyramidBarClustered=109,
        xlPyramidBarStacked=110,
        xlPyramidBarStacked100=111,
        xlPyramidCol=112,
        xlPyramidColClustered=106,
        xlPyramidColStacked=107,
        xlPyramidColStacked100=108,
        xlRadar=-4151,
        xlRadarFilled=82,
        xlRadarMarkers=81,
        xlStockHLC=88,
        xlStockOHLC=89,
        xlStockVHLC=90,
        xlStockVOHLC=91,
        xlSurface=83,
        xlSurfaceTopView=85,
        xlSurfaceTopViewWireframe=86,
        xlSurfaceWireframe=84,
        xlXYScatter=-4169,
        xlXYScatterLines=74,
        xlXYScatterLinesNoMarkers=75,
        xlXYScatterSmooth=72,
        xlXYScatterSmoothNoMarkers=73
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
           ChartType = type;
           CreateNewChart(drawings, type);
           _chartPath = rootPath + "/" + GetChartNodeText();
           WorkSheet = drawings.Worksheet;

           string chartNodeText=GetChartNodeText();
           _groupingPath = string.Format(_groupingPath, chartNodeText);

           _chartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, ChartXml.SelectSingleNode(_chartPath, drawings.NameSpaceManager));
           _chartXmlHelper = new XmlHelper(drawings.NameSpaceManager, ChartXml);

           SetTypeProperties(drawings);
       }
       public ExcelWorksheet WorkSheet { get; internal set; }

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
           Part = package.CreatePart(UriChart, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", CompressionOption.Maximum);

           StreamWriter streamChart = new StreamWriter(Part.GetStream(FileMode.Create, FileAccess.Write));
           ChartXml.Save(streamChart);
           streamChart.Close();
           package.Flush();

           PackageRelationship chartRelation = drawings.Part.CreateRelationship(UriChart, TargetMode.Internal, ExcelPackage.schemaRelationships + "/chart");
           TopNode.SelectSingleNode("//c:chart", NameSpaceManager).Attributes["r:id"].Value = chartRelation.Id;
           package.Flush();
       }

       private string ChartStartXml(eChartType type)
       {
           StringBuilder xml=new StringBuilder();
           int axID=1;
           int xAxID=2;
           xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
           xml.AppendFormat("<c:chartSpace xmlns:c=\"{0}\" xmlns:a=\"{1}\" xmlns:r=\"{2}\">", ExcelPackage.schemaChart, ExcelPackage.schemaMain, ExcelPackage.schemaRelationships);
           xml.Append("<c:date1904 val=\"1\"/><c:lang val=\"sv-SE\"/><c:chart>");
           xml.AppendFormat("{0}<c:plotArea><c:layout/>",AddPerspectiveXml(type));

           xml.AppendFormat("<{0}>{6}{3}{9}{10}{4}{5}{7}{8}<c:axId val=\"{1}\"/><c:axId val=\"{2}\"/></{0}>", GetChartNodeText(), axID, xAxID, AddBarDir(type), AddMarker(type), AddShape(type), AddVaryColors(), AddFirstSliceAng(type), AddHoleSize(type), AddScatterType(type), AddGrouping());

           xml.AppendFormat("<c:catAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx><c:valAx><c:axId val=\"{1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:axPos val=\"l\"/><c:majorGridlines/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx></c:plotArea><c:legend><c:legendPos val=\"r\"/><c:layout/></c:legend><c:plotVisOnly val=\"1\"/></c:chart>", axID, xAxID);
           xml.Append("<c:printSettings><c:headerFooter/><c:pageMargins b=\"0.75\" l=\"0.7\" r=\"0.7\" t=\"0.75\" header=\"0.3\" footer=\"0.3\"/><c:pageSetup/></c:printSettings></c:chartSpace>");          
           return xml.ToString();
       }

       private string AddScatterType(eChartType type)
       {
           if (type == eChartType.xlXYScatter ||
               type == eChartType.xlXYScatterLines ||
               type == eChartType.xlXYScatterLinesNoMarkers ||
               type == eChartType.xlXYScatterSmooth ||
               type == eChartType.xlXYScatterSmoothNoMarkers)
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
           if(IsTypeClustered() || IsTypePercentStacked() || IsTypeStacked() || IsTypePieDoughnut())
           {
               return "<c:grouping val=\"standard\"/>";
           }
           else
           {
               return "";
           }
       }
       private object AddHoleSize(eChartType type)
       {
           if (type == eChartType.xlDoughnut ||
               type == eChartType.xlDoughnutExploded)
           {
               return "<c:holeSize val=\"50\" />";
           }
           else
           {
               return "";
           }
       }

       private object AddFirstSliceAng(eChartType type)
       {
           if (type == eChartType.xlDoughnut ||
               type == eChartType.xlDoughnutExploded)
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
       private string AddMarker(eChartType type)
       {
           if (type == eChartType.xlLineMarkers ||
               type == eChartType.xlLineMarkersStacked ||
               type == eChartType.xlLineMarkersStacked100 ||
               type == eChartType.xlXYScatterLines ||
               type == eChartType.xlXYScatterSmooth)
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

       private object AddBarDir(eChartType type)
       {
 	        if( type == eChartType.xl3DBarClustered ||
                type == eChartType.xl3DBarStacked ||
                type == eChartType.xl3DBarStacked100 ||
                type == eChartType.xlBarClustered || 
                type == eChartType.xlBarStacked || 
                type == eChartType.xlBarStacked100 || 
                type == eChartType.xlBarOfPie)
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

        protected bool IsType3D()
        {
            return ChartType == eChartType.xl3DArea ||
                            ChartType == eChartType.xl3DAreaStacked ||
                            ChartType == eChartType.xl3DAreaStacked100 ||
                            ChartType == eChartType.xl3DBarClustered ||
                            ChartType == eChartType.xl3DBarStacked ||
                            ChartType == eChartType.xl3DBarStacked100 ||
                            ChartType == eChartType.xl3DColumn ||
                            ChartType == eChartType.xl3DColumnClustered ||
                            ChartType == eChartType.xl3DColumnStacked ||
                            ChartType == eChartType.xl3DColumnStacked100 ||
                            ChartType == eChartType.xl3DLine ||
                            ChartType == eChartType.xl3DPie ||
                            ChartType == eChartType.xl3DPieExploded ||
                            ChartType == eChartType.xlBubble3DEffect ||
                            ChartType == eChartType.xlConeBarClustered ||
                            ChartType == eChartType.xlConeBarStacked ||
                            ChartType == eChartType.xlConeBarStacked100 ||
                            ChartType == eChartType.xlConeCol ||
                            ChartType == eChartType.xlConeColClustered ||
                            ChartType == eChartType.xlConeColStacked ||
                            ChartType == eChartType.xlConeColStacked100 ||
                            ChartType == eChartType.xlCylinderBarClustered ||
                            ChartType == eChartType.xlCylinderBarStacked ||
                            ChartType == eChartType.xlCylinderBarStacked100 ||
                            ChartType == eChartType.xlCylinderCol ||
                            ChartType == eChartType.xlCylinderColClustered ||
                            ChartType == eChartType.xlCylinderColStacked ||
                            ChartType == eChartType.xlCylinderColStacked100 ||
                            ChartType == eChartType.xlPyramidBarClustered ||
                            ChartType == eChartType.xlPyramidBarStacked ||
                            ChartType == eChartType.xlPyramidBarStacked100 ||
                            ChartType == eChartType.xlPyramidCol ||
                            ChartType == eChartType.xlPyramidColClustered ||
                            ChartType == eChartType.xlPyramidColStacked ||
                            ChartType == eChartType.xlPyramidColStacked100 ||
                            ChartType == eChartType.xlDoughnut ||
                            ChartType == eChartType.xlDoughnutExploded;
        }
        protected bool IsTypeShape()
        {
            return ChartType == eChartType.xl3DBarClustered ||
                    ChartType == eChartType.xl3DBarStacked ||
                    ChartType == eChartType.xl3DBarStacked100 ||
                    ChartType == eChartType.xl3DBarClustered ||
                    ChartType == eChartType.xl3DBarStacked ||
                    ChartType == eChartType.xl3DBarStacked100 ||
                    ChartType == eChartType.xl3DColumn ||
                    ChartType == eChartType.xl3DColumnClustered ||
                    ChartType == eChartType.xl3DColumnStacked ||
                    ChartType == eChartType.xl3DColumnStacked100 ||
                //ChartType == eChartType.xl3DPie ||
                //ChartType == eChartType.xl3DPieExploded ||
                    ChartType == eChartType.xlBubble3DEffect ||
                    ChartType == eChartType.xlConeBarClustered ||
                    ChartType == eChartType.xlConeBarStacked ||
                    ChartType == eChartType.xlConeBarStacked100 ||
                    ChartType == eChartType.xlConeCol ||
                    ChartType == eChartType.xlConeColClustered ||
                    ChartType == eChartType.xlConeColStacked ||
                    ChartType == eChartType.xlConeColStacked100 ||
                    ChartType == eChartType.xlCylinderBarClustered ||
                    ChartType == eChartType.xlCylinderBarStacked ||
                    ChartType == eChartType.xlCylinderBarStacked100 ||
                    ChartType == eChartType.xlCylinderCol ||
                    ChartType == eChartType.xlCylinderColClustered ||
                    ChartType == eChartType.xlCylinderColStacked ||
                    ChartType == eChartType.xlCylinderColStacked100 ||
                    ChartType == eChartType.xlPyramidBarClustered ||
                    ChartType == eChartType.xlPyramidBarStacked ||
                    ChartType == eChartType.xlPyramidBarStacked100 ||
                    ChartType == eChartType.xlPyramidCol ||
                    ChartType == eChartType.xlPyramidColClustered ||
                    ChartType == eChartType.xlPyramidColStacked ||
                    ChartType == eChartType.xlPyramidColStacked100; //||
                    //ChartType == eChartType.xlDoughnut ||
                    //ChartType == eChartType.xlDoughnutExploded;
        }
        protected bool IsTypePercentStacked()
        {
            return ChartType == eChartType.xlAreaStacked100 ||
                           ChartType == eChartType.xlBarStacked100 ||
                           ChartType == eChartType.xlConeBarStacked100 ||
                           ChartType == eChartType.xlConeColStacked100 ||
                           ChartType == eChartType.xlCylinderBarStacked100 ||
                           ChartType == eChartType.xlCylinderColStacked ||
                           ChartType == eChartType.xlLineMarkersStacked100 ||
                           ChartType == eChartType.xlLineStacked100 ||
                           ChartType == eChartType.xlPyramidBarStacked100 ||
                           ChartType == eChartType.xlPyramidColStacked100;
        }
        protected bool IsTypeStacked()
        {
            return ChartType == eChartType.xlAreaStacked ||
                           ChartType == eChartType.xlBarStacked ||
                           ChartType == eChartType.xl3DColumnStacked ||
                           ChartType == eChartType.xlConeBarStacked ||
                           ChartType == eChartType.xlConeColStacked ||
                           ChartType == eChartType.xlCylinderBarStacked ||
                           ChartType == eChartType.xlCylinderColStacked ||
                           ChartType == eChartType.xlLineMarkersStacked ||
                           ChartType == eChartType.xlLineStacked ||
                           ChartType == eChartType.xlPyramidBarStacked ||
                           ChartType == eChartType.xlPyramidColStacked;
        }
        protected bool IsTypeClustered()
        {
            return ChartType == eChartType.xlBarClustered ||
                           ChartType == eChartType.xl3DBarClustered ||
                           ChartType == eChartType.xl3DColumnClustered ||
                           ChartType == eChartType.xlColumnClustered ||
                           ChartType == eChartType.xlConeBarClustered ||
                           ChartType == eChartType.xlConeColClustered ||
                           ChartType == eChartType.xlCylinderBarClustered ||
                           ChartType == eChartType.xlCylinderColClustered ||
                           ChartType == eChartType.xlPyramidBarClustered ||
                           ChartType == eChartType.xlPyramidColClustered;
        }
        protected bool IsTypePieDoughnut()
        {
            return ChartType == eChartType.xlPie ||
                           ChartType == eChartType.xlPieExploded ||
                           ChartType == eChartType.xlPieOfPie ||
                           ChartType == eChartType.xl3DPie ||
                           ChartType == eChartType.xl3DPieExploded ||
                           ChartType == eChartType.xlBarOfPie ||
                           ChartType == eChartType.xlDoughnut ||
                           ChartType == eChartType.xlDoughnutExploded;
        }
        private void LoadAxis()
        {
            XmlNodeList nl = ChartXml.SelectNodes(_chartPath + "/c:axId", NameSpaceManager);
            List<ExcelChartAxis> l=new List<ExcelChartAxis>();
            foreach (XmlNode node in nl)
            {
                string id = node.Attributes["val"].Value;
                XmlNode axisNode = ChartXml.SelectNodes(rootPath + string.Format("/*/c:axId[@val=\"{0}\"]",id), NameSpaceManager)[1].ParentNode;
                ExcelChartAxis ax = new ExcelChartAxis(NameSpaceManager, axisNode);
                l.Add(ax);
            }
            _axis = l.ToArray();
        }
        private void SetChartType()
        {
            ChartType = 0;
            foreach(XmlNode n in ChartXml.SelectSingleNode(rootPath, _drawings.NameSpaceManager).ChildNodes)
            {
                switch(n.Name)
                {
                    case "c:area3DChart":
                        ChartType = eChartType.xl3DArea;
                        break;
                    case "c:areaChart":
                        ChartType = eChartType.xlArea;
                        break;
                    case "c:barChart":
                        ChartType = eChartType.xlBarClustered;
                        break;
                    case "c:bar3DChart":
                        ChartType = eChartType.xl3DBarClustered;
                        break;
                    case "c:bubbleChart":
                        ChartType = eChartType.xlBubble;
                        break;
                    case "c:doughnutChart":
                        ChartType = eChartType.xlDoughnut;
                        break;
                    case "c:lineChart":
                        ChartType = eChartType.xlLine;
                        break;
                    case "c:line3DChart":
                        ChartType = eChartType.xl3DLine;
                        break;
                    case "c:pie3DChart":
                        ChartType = eChartType.xl3DPie;
                        break;
                    case "c:pieChart":
                        ChartType = eChartType.xlPie;
                        break;
                    case "c:radarChart":
                        ChartType = eChartType.xlRadar;
                        break;
                    case "c:scatterChart":
                        ChartType = eChartType.xlXYScatter;
                        break;
                    case "c:surface3DChart":
                    case "c:surfaceChart":
                        ChartType = eChartType.xlSurface;
                        break;
                    case "c:stockChart":
                        ChartType = eChartType.xlStockHLC;
                        break;
                }
                if (ChartType != 0)
                {
                    _chartPath = rootPath + "/" + n.Name;
                    return;
                }
            }
        }

        protected string GetChartNodeText()
        {
            switch (ChartType)
            {
                case eChartType.xl3DArea:
                case eChartType.xl3DAreaStacked:
                case eChartType.xl3DAreaStacked100:
                    return "c:area3DChart";
                case eChartType.xlArea:
                case eChartType.xlAreaStacked:
                case eChartType.xlAreaStacked100:
                    return "c:areaChart";
                case eChartType.xlBarClustered:
                case eChartType.xlBarStacked:
                case eChartType.xlBarStacked100:
                    return "c:barChart";
                case eChartType.xl3DBarClustered:
                case eChartType.xl3DBarStacked:
                case eChartType.xl3DBarStacked100:
                case eChartType.xl3DColumnClustered:
                case eChartType.xl3DColumnStacked:
                case eChartType.xl3DColumnStacked100:
                case eChartType.xlConeBarClustered:
                case eChartType.xlConeBarStacked:
                case eChartType.xlConeBarStacked100:
                case eChartType.xlConeCol:
                case eChartType.xlConeColClustered:
                case eChartType.xlConeColStacked:
                case eChartType.xlConeColStacked100:
                case eChartType.xlCylinderBarClustered:
                case eChartType.xlCylinderBarStacked:
                case eChartType.xlCylinderBarStacked100:
                case eChartType.xlCylinderCol:
                case eChartType.xlCylinderColClustered:
                case eChartType.xlCylinderColStacked:
                case eChartType.xlCylinderColStacked100:
                case eChartType.xlPyramidBarClustered:
                case eChartType.xlPyramidBarStacked:
                case eChartType.xlPyramidBarStacked100:
                case eChartType.xlPyramidCol:
                case eChartType.xlPyramidColClustered:
                case eChartType.xlPyramidColStacked:
                case eChartType.xlPyramidColStacked100:
                    return "c:bar3DChart";
                case eChartType.xlBubble:
                    return "c:bubbleChart";
                case eChartType.xlDoughnut:
                case eChartType.xlDoughnutExploded:
                    return "c:doughnutChart";
                case eChartType.xlLine:
                case eChartType.xlLineMarkers:
                case eChartType.xlLineMarkersStacked:
                case eChartType.xlLineMarkersStacked100:
                case eChartType.xlLineStacked:
                case eChartType.xlLineStacked100:
                    return "c:lineChart";
                case eChartType.xl3DLine:
                    return "c:line3DChart";
                case eChartType.xlPie:
                case eChartType.xlPieExploded:
                    return "c:pieChart";
                case eChartType.xlBarOfPie:
                case eChartType.xlPieOfPie:
                    return "c:ofPieChart";
                case eChartType.xl3DPie:
                case eChartType.xl3DPieExploded:
                    return "c:pie3DChart";
                case eChartType.xlRadar:
                case eChartType.xlRadarFilled:
                case eChartType.xlRadarMarkers:
                    return "c:radarChart";
                case eChartType.xlXYScatter:
                case eChartType.xlXYScatterLines:
                case eChartType.xlXYScatterLinesNoMarkers:
                case eChartType.xlXYScatterSmooth:
                case eChartType.xlXYScatterSmoothNoMarkers:
                    return "c:scatterChart";
                case eChartType.xlSurface:
                    return "c:surfaceChart";
                case eChartType.xlStockHLC:
                    return "c:stockChart";
                default:
                    throw(new NotImplementedException("Chart type not implemented"));
            }
        }
        internal PackagePart Part { get; set; }
        public XmlDocument ChartXml { get; set; }
        internal Uri UriChart { get; set; }
        public eChartType ChartType { get; set; }
        public ExcelChartSeries Series
        {
            get
            {
                return _chartSeries;
            }
        }
        public ExcelChartAxis[] Axis
        {
            get
            {
                return _axis;
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
       const string titlePath = "c:chartSpace/c:chart/c:title/c:tx/c:rich/a:p/a:r/a:t";
       public string Header
       {
           get
           {
               return _chartXmlHelper.GetXmlNode(titlePath);
           }
           set
           {
               _chartXmlHelper.CreateNode(titlePath);
               _chartXmlHelper.SetXmlNode(titlePath, value);
           }
       }
       string _groupingPath = "c:chartSpace/c:chart/c:plotArea/{0}/c:grouping/@val";
       public eGrouping Grouping
       {
           get
           {
               return GetGroupingEnum(_chartXmlHelper.GetXmlNode(_groupingPath));
           }
           set
           {
               _chartXmlHelper.SetXmlNode(_groupingPath, GetGroupingText(value));
           }
       }
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
       internal string Id
        {
            get { return ""; }
        }
    }
}
