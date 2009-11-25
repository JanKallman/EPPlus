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

namespace OfficeOpenXml.Drawing
{
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
   /// <summary>
   /// Provide access to Chart objects.
   /// </summary>
    public class ExcelChart : ExcelDrawing
    {
       const string rootPath = "c:chartSpace/c:chart/c:plotArea";
       string _chartPath;
       ExcelChartSeries _series;
       ExcelChartAxis[] _axis;
       XmlHelper _chartXmlHelper;       
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
                SetChartType();
                _chartXmlHelper = new XmlHelper(drawings.NameSpaceManager, ChartXml);
                _series = new ExcelChartSeries(this, drawings.NameSpaceManager, ChartXml.SelectSingleNode(_chartPath, drawings.NameSpaceManager));
                LoadAxis();
            }
            else
            {
                ChartXml = null;
            }
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
        internal PackagePart Part { get; set; }
        public XmlDocument ChartXml { get; set; }
        internal Uri UriChart { get; set; }
        public eChartType ChartType { get; set; }
        public ExcelChartSeries Series
        {
            get
            {
                return _series;
            }
        }
        public ExcelChartAxis[] Axis
        {
            get
            {
                return _axis;
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
       internal string Id
        {
            get { return ""; }
        }
    }
}
