using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    public class ExcelLineChart  : ExcelChart
    {
        #region "Constructors"
        internal ExcelLineChart(ExcelDrawings drawings, XmlNode node) :
            base(drawings, node)
        {
            _chartTopPath = string.Format(_chartTopPath, GetChartNodeText());
        }
        internal ExcelLineChart(ExcelDrawings drawings, XmlNode node, eChartType type) :
            base(drawings, node, type)
        {
            _chartTopPath = string.Format(_chartTopPath, GetChartNodeText());
        }
        #endregion
        string _chartTopPath = "c:chartSpace/c:chart/c:plotArea/{0}";
        ExcelChartDataLabel _DataLabel = null;
        public ExcelChartDataLabel DataLabel
        {
            get
            {
                if (_DataLabel == null)
                {
                    _DataLabel = new ExcelChartDataLabel(NameSpaceManager, _chartXmlHelper.TopNode.SelectSingleNode(_chartTopPath, NameSpaceManager));
                }
                return _DataLabel;
            }
        }    
    }
}
