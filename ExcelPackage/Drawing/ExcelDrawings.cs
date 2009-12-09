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
using System.IO.Packaging;
using System.Xml;
using System.Collections;
using System.IO;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Enables access to the Drawings.
    /// VERY basic support for Charts, Shapes and Pictures in present version.
    /// </summary>
    public class ExcelDrawings : IEnumerable
    {
        private XmlDocument _drawingsXml=new XmlDocument();
        private Dictionary<int, ExcelDrawing> _drawings;
        public ExcelDrawings(ExcelPackage xlPackage, ExcelWorksheet sheet)
        {
                _drawingsXml = new XmlDocument();
                _drawingsXml.PreserveWhitespace = true;
                _drawings = new Dictionary<int, ExcelDrawing>();
                Worksheet = sheet;
                XmlNode node = sheet.WorksheetXml.SelectSingleNode("//d:drawing", sheet.NameSpaceManager);
                CreateNSM();
                if (node != null)
                {
                    PackageRelationship drawingRelation = sheet.Part.GetRelationship(node.Attributes["r:id"].Value);
                    _uriDrawing = PackUriHelper.ResolvePartUri(sheet.WorksheetUri, drawingRelation.TargetUri);

                    _part = xlPackage.Package.GetPart(_uriDrawing);
                    _drawingsXml.Load(_part.GetStream());

                    AddDrawings();
                }
         }
        internal ExcelWorksheet Worksheet { get; set; }
        public XmlDocument DrawingXml
        {
            get
            {
                return _drawingsXml;
            }
        }
        private void AddDrawings()
        {
            XmlNodeList list = _drawingsXml.SelectNodes("//xdr:twoCellAnchor", NameSpaceManager);

            int i = 0;
            foreach (XmlNode node in list)
            {
                ExcelDrawing dr = ExcelDrawing.GetDrawing(this, node);
                _drawings.Add(++i, dr);
            }
        }


        #region NamespaceManager
        /// <summary>
        /// Creates the NamespaceManager. 
        /// </summary>
        private void CreateNSM()
        {
            NameTable nt = new NameTable();
            _nsManager = new XmlNamespaceManager(nt);
            _nsManager.AddNamespace("a", ExcelPackage.schemaDrawings);
            _nsManager.AddNamespace("xdr", ExcelPackage.schemaSheetDrawings);
            _nsManager.AddNamespace("c", ExcelPackage.schemaChart);
            _nsManager.AddNamespace("r", ExcelPackage.schemaRelationships);
        }
        /// <summary>
        /// Provides access to a namespace manager instance to allow XPath searching
        /// </summary>
        XmlNamespaceManager _nsManager=null;
        public XmlNamespaceManager NameSpaceManager
        {
            get
            {
                return _nsManager;
            }
        }
        #endregion
        #region IEnumerable Members

        public IEnumerator GetEnumerator()
        {
            return (_drawings.Values.GetEnumerator());
        }
        /// <summary>
        /// Returns the worksheet at the specified position.  
        /// </summary>
        /// <param name="PositionID">The position of the worksheet. 1-base</param>
        /// <returns></returns>
        public ExcelDrawing this[int PositionID]
        {
            get
            {
                return (_drawings[PositionID]);
            }
        }

        /// <summary>
        /// Returns the worksheet matching the specified name
        /// </summary>
        /// <param name="Name">The name of the worksheet</param>
        /// <returns></returns>
        public ExcelDrawing this[string Name]
        {
            get
            {
                foreach (ExcelDrawing drawing in _drawings.Values)
                {
                    if (drawing.Name == Name)
                        return drawing;
                }
                return null;
                //throw new Exception(string.Format("ExcelWorksheets Error: Worksheet '{0}' not found!",Name));
            }
        }
        public int Count
        {
            get
            {
                if (_drawings == null)
                {
                    return 0;
                }
                else
                {
                    return _drawings.Count;
                }
            }
        }
        PackagePart _part=null;
        internal PackagePart Part
        {
            get
            {
                return _part;
            }        
        }
        Uri _uriDrawing=null;
        public Uri UriDrawing
        {
            get
            {
                return _uriDrawing;
            }
        }
        #endregion
        #region "Add functions"
            public ExcelChart AddChart(string Name, eChartType ChartType)
            {
                if (ChartType == eChartType.xlBubble ||
                    ChartType == eChartType.xlBubble3DEffect ||
                    ChartType == eChartType.xlRadar ||
                    ChartType == eChartType.xlRadarFilled ||
                    ChartType == eChartType.xlRadarMarkers ||
                    ChartType == eChartType.xlStockHLC ||
                    ChartType == eChartType.xlStockOHLC ||
                    ChartType == eChartType.xlStockVOHLC ||
                    ChartType == eChartType.xlSurface ||
                    ChartType == eChartType.xlSurfaceTopView ||
                    ChartType == eChartType.xlSurfaceTopViewWireframe ||
                    ChartType == eChartType.xlSurfaceWireframe)
                {
                    throw(new NotImplementedException("Chart type not supported in this version"));
                }

                XmlElement drawNode = CreateDrawingXml();

                
                ExcelChart chart = GetNewChart(drawNode, ChartType);
                SetDrawingDefaults(chart, Name);
                _drawings.Add(_drawings.Count + 1, chart);
                return chart;
            }

            private ExcelChart GetNewChart(XmlNode drawNode, eChartType chartType)
            {
                switch(chartType)
                {
                    case eChartType.xlPie:
                    case eChartType.xlPieExploded:
                    case eChartType.xl3DPie:
                    case eChartType.xl3DPieExploded:
                        return new ExcelPieChart(this, drawNode, chartType);
                    case eChartType.xlBarOfPie:
                    case eChartType.xlPieOfPie:
                        return new ExcelOfPieChart(this, drawNode, chartType);
                    case eChartType.xlDoughnut:
                    case eChartType.xlDoughnutExploded:
                        return new ExcelDoughnutChart(this, drawNode, chartType);
                    case eChartType.xlBarClustered:
                    case eChartType.xlBarStacked:
                    case eChartType.xlBarStacked100:
                    case eChartType.xl3DBarClustered:
                    case eChartType.xl3DBarStacked:
                    case eChartType.xl3DBarStacked100:
                    case eChartType.xlConeBarClustered:
                    case eChartType.xlConeBarStacked:
                    case eChartType.xlConeBarStacked100:
                    case eChartType.xlCylinderBarClustered:
                    case eChartType.xlCylinderBarStacked:
                    case eChartType.xlCylinderBarStacked100:
                    case eChartType.xlPyramidBarClustered:
                    case eChartType.xlPyramidBarStacked:
                    case eChartType.xlPyramidBarStacked100:
                    case eChartType.xlColumnClustered:
                    case eChartType.xlColumnStacked:
                    case eChartType.xlColumnStacked100:
                    case eChartType.xl3DColumn:
                    case eChartType.xl3DColumnClustered:
                    case eChartType.xl3DColumnStacked:
                    case eChartType.xl3DColumnStacked100:
                    case eChartType.xlConeCol:
                    case eChartType.xlConeColClustered:
                    case eChartType.xlConeColStacked:
                    case eChartType.xlConeColStacked100:
                    case eChartType.xlCylinderCol:
                    case eChartType.xlCylinderColClustered:
                    case eChartType.xlCylinderColStacked:
                    case eChartType.xlCylinderColStacked100:
                    case eChartType.xlPyramidCol:
                    case eChartType.xlPyramidColClustered:
                    case eChartType.xlPyramidColStacked:
                    case eChartType.xlPyramidColStacked100:
                        return new ExcelBarChart(this, drawNode, chartType);
                    case eChartType.xlXYScatter:
                    case eChartType.xlXYScatterLines:
                    case eChartType.xlXYScatterLinesNoMarkers:
                    case eChartType.xlXYScatterSmooth:
                    case eChartType.xlXYScatterSmoothNoMarkers:
                        return new ExcelScatterChart(this, drawNode, chartType);
                    default:
                        return new ExcelChart(this, drawNode, chartType);
                }
            }

            private void SetDrawingDefaults(ExcelChart chart, string Name/*, int fromRow, int fromColumn, int toRow, int toColumn*/)
            {
                chart.Name = Name;  
                chart.From.Column = 0;
                chart.From.ColumnOff = 0;
                chart.From.Row = 0;
                chart.From.RowOff = 0;

                chart.To.Column = 10;
                chart.To.ColumnOff = 0;
                chart.To.Row = 10;
                chart.To.RowOff = 0;
            }

            private XmlElement CreateDrawingXml()
            {
                if (DrawingXml.OuterXml == "")
                {
                    DrawingXml.LoadXml(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><xdr:wsDr xmlns:xdr=\"{0}\" xmlns:a=\"{1}\" />", ExcelPackage.schemaSheetDrawings, ExcelPackage.schemaDrawings));
                    _uriDrawing = new Uri(string.Format("/xl/drawings/drawing{0}.xml", Worksheet.SheetID),UriKind.Relative);

                    Package package = Worksheet.xlPackage.Package;
                    _part = package.CreatePart(_uriDrawing, "application/vnd.openxmlformats-officedocument.drawing+xml", CompressionOption.Maximum);

                    StreamWriter streamChart = new StreamWriter(_part.GetStream(FileMode.Create, FileAccess.Write));
                    DrawingXml.Save(streamChart);
                    streamChart.Close();
                    package.Flush();

                    PackageRelationship drawRelation = Worksheet.Part.CreateRelationship(_uriDrawing, TargetMode.Internal, ExcelPackage.schemaRelationships + "/drawing");
                    XmlElement e = Worksheet.WorksheetXml.CreateElement("drawing", ExcelPackage.schemaMain);
                    e.SetAttribute("id",ExcelPackage.schemaRelationships, drawRelation.Id);

                    Worksheet.WorksheetXml.DocumentElement.AppendChild(e);
                    package.Flush();                    
                }
                XmlNode colNode = _drawingsXml.SelectSingleNode("//xdr:wsDr", NameSpaceManager);
                XmlElement drawNode = _drawingsXml.CreateElement("twoCellAnchor", ExcelPackage.schemaSheetDrawings);
                colNode.AppendChild(drawNode);

                //Add from position Element;
                XmlElement fromNode = _drawingsXml.CreateElement("from", ExcelPackage.schemaSheetDrawings);
                drawNode.AppendChild(fromNode);

                //Add to position Element;
                XmlElement toNode = _drawingsXml.CreateElement("to", ExcelPackage.schemaSheetDrawings);
                drawNode.AppendChild(toNode);
                return drawNode;
            }
        #endregion
    }
}
