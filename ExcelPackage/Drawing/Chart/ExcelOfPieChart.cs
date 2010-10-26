using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO.Packaging;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Provides access to pie chart specific properties
    /// </summary>
    public class ExcelOfPieChart : ExcelPieChart
    {
        //internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node) :
        //    base(drawings, node)
        //{

        //}
        internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node, eChartType type) :
            base(drawings, node, type)
        {
                SetTypeProperties();
        }
        internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart) :
            base(drawings, node, type, topChart)
        {
            SetTypeProperties();
        }

        public ExcelOfPieChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, PackagePart part, XmlDocument chartXml, XmlNode chartNode) :
           base(drawings, node, uriChart, part, chartXml, chartNode)
        {
            SetTypeProperties();
        }

        private void SetTypeProperties()
        {
            if (ChartType == eChartType.BarOfPie)
            {
                OfPieType = ePieType.Bar;
            }
            else
            {
                OfPieType = ePieType.Pie;
            }
        }

        const string pieTypePath = "c:chartSpace/c:chart/c:plotArea/c:ofPieChart/c:ofPieType/@val";
        private ExcelDrawings drawings;
        private XmlNode node;
        public ePieType OfPieType
        {
            get
            {
                if (_chartXmlHelper.GetXmlNodeString(pieTypePath) == "bar")
                    return ePieType.Bar;
                else
                {
                    return ePieType.Pie;
                }
            }
            internal set
            {
                _chartXmlHelper.CreateNode(pieTypePath,true);
                _chartXmlHelper.SetXmlNodeString(pieTypePath, value == ePieType.Bar ? "bar" : "pie");
            }
        }

    }
}
