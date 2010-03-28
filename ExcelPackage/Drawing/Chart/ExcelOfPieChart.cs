using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Provides access to pie chart specific properties
    /// </summary>
    public class ExcelOfPieChart : ExcelPieChart
    {
        internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node) :
            base(drawings, node)
        {

        }
        internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node, eChartType type) :
            base(drawings, node, type)
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
        public ePieType OfPieType
        {
            get
            {
                if (_chartXmlHelper.GetXmlNode(pieTypePath) == "bar")
                    return ePieType.Bar;
                else
                {
                    return ePieType.Pie;
                }
            }
            internal set
            {
                _chartXmlHelper.CreateNode(pieTypePath,true);
                _chartXmlHelper.SetXmlNode(pieTypePath, value == ePieType.Bar ? "bar" : "pie");
            }
        }

    }
}
