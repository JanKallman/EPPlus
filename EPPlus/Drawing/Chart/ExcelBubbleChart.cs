using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Provides access to bubble chart specific properties
    /// </summary>
    public sealed class ExcelBubbleChart : ExcelChart
    {
        internal ExcelBubbleChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
            base(drawings, node, type, topChart, PivotTableSource)
        {
            ShowNegativeBubbles = false;
            BubbleScale = 100;
            _chartSeries = new ExcelBubbleChartSeries(this, drawings.NameSpaceManager, _chartNode, PivotTableSource!=null);
            //SetTypeProperties();
        }

        internal ExcelBubbleChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot) :
            base(drawings, node, type, isPivot)
        {
            _chartSeries = new ExcelBubbleChartSeries(this, drawings.NameSpaceManager, _chartNode, isPivot);
            //SetTypeProperties();
        }
        internal ExcelBubbleChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
            base(drawings, node, uriChart, part, chartXml, chartNode)
        {
            _chartSeries = new ExcelBubbleChartSeries(this, _drawings.NameSpaceManager, _chartNode, false);
            //SetTypeProperties();
        }
        internal ExcelBubbleChart(ExcelChart topChart, XmlNode chartNode) :
            base(topChart, chartNode)
        {
            _chartSeries = new ExcelBubbleChartSeries(this, _drawings.NameSpaceManager, _chartNode, false);
        }
        string BUBBLESCALE_PATH = "c:bubbleScale/@val";
        /// <summary>
        /// Specifies the scale factor for the bubble chart. Can range from 0 to 300, corresponding to a percentage of the default size,
        /// </summary>
        public int BubbleScale
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeInt(BUBBLESCALE_PATH);
            }
            set
            {
                if(value < 0 && value > 300)
                {
                    throw(new ArgumentOutOfRangeException("Bubblescale out of range. 0-300 allowed"));
                }
                _chartXmlHelper.SetXmlNodeString(BUBBLESCALE_PATH, value.ToString());
            }
        }
        string SHOWNEGBUBBLES_PATH = "c:showNegBubbles/@val";
        /// <summary>
        /// Specifies the scale factor for the bubble chart. Can range from 0 to 300, corresponding to a percentage of the default size,
        /// </summary>
        public bool ShowNegativeBubbles
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeBool(SHOWNEGBUBBLES_PATH);
            }
            set
            {
                _chartXmlHelper.SetXmlNodeBool(BUBBLESCALE_PATH, value, true);
            }
        }
        string BUBBLE3D_PATH = "c:bubble3D/@val";
        /// <summary>
        /// Specifies if the bubblechart is three dimensional
        /// </summary>
        public bool Bubble3D
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeBool(BUBBLE3D_PATH);
            }
            set
            {
                _chartXmlHelper.SetXmlNodeBool(BUBBLE3D_PATH, value);
                ChartType = value ? eChartType.Bubble3DEffect : eChartType.Bubble;
            }
        }
        string SIZEREPRESENTS_PATH = "c:sizeRepresents/@val";
        /// <summary>
        /// Specifies the scale factor for the bubble chart. Can range from 0 to 300, corresponding to a percentage of the default size,
        /// </summary>
        public eSizeRepresents SizeRepresents
        {
            get
            {
                var v = _chartXmlHelper.GetXmlNodeString(SIZEREPRESENTS_PATH).ToLower();
                if (v == "w")
                {
                    return eSizeRepresents.Width;
                }
                else
                {
                    return eSizeRepresents.Area;
                }
            }
            set
            {
                _chartXmlHelper.SetXmlNodeString(SIZEREPRESENTS_PATH, value == eSizeRepresents.Width ? "w" : "area");
            }
        }
        public new ExcelBubbleChartSeries Series
        {
            get
            {

                return (ExcelBubbleChartSeries)_chartSeries;
            }
        }
        internal override eChartType GetChartType(string name)
        {
            if (Bubble3D)
            {
                return eChartType.Bubble3DEffect;
            }
            else
            {
                return eChartType.Bubble;
            }
        }
    }
}
