using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO.Packaging;

namespace OfficeOpenXml.Drawing.Chart
{
    public class ExcelLineChart  : ExcelChart
    {
        #region "Constructors"
        //internal ExcelLineChart(ExcelDrawings drawings, XmlNode node) :
        //    base(drawings, node)
        //{
        //    _chartTopPath = string.Format(_chartTopPath, GetChartNodeText());
        //}
        internal ExcelLineChart(ExcelDrawings drawings, XmlNode node, eChartType type) :
            base(drawings, node, type)
        {
            //_chartTopPath = string.Format(_chartTopPath, GetChartNodeText());
        }

        internal ExcelLineChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, PackagePart part, XmlDocument chartXml, XmlNode chartNode) :
            base(drawings, node, uriChart, part, chartXml, chartNode)
        {
            //_chartTopPath = string.Format(_chartTopPath, chartNode.Name);
        }

        internal ExcelLineChart(ExcelChart topChart, XmlNode chartNode) :
            base(topChart, chartNode)
        {
           // _chartTopPath = string.Format(_chartTopPath, chartNode.Name);
        }
        internal ExcelLineChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart) :
            base(drawings, node, type, topChart)
        {
            //_chartTopPath = string.Format(_chartTopPath, GetChartNodeText());
        }
        #endregion
        //string _chartTopPath = "c:chartSpace/c:chart/c:plotArea/{0}";
        ExcelChartDataLabel _DataLabel = null;
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
    }
}
