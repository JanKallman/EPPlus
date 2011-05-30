using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Enumerates charttypes 
    /// </summary>
    public class ExcelChartCollection : IEnumerable<ExcelChart>
    {
        List<ExcelChart> _list = new List<ExcelChart>();
        ExcelChart _topChart;
        internal ExcelChartCollection(ExcelChart chart)
        {
            _topChart = chart;
            _list.Add(chart);
        }
        internal void Add(ExcelChart chart)
        {
            _list.Add(chart);
        }
        /// <summary>
        /// Add a new charttype to the chart
        /// </summary>
        /// <param name="chartType">The type of the new chart</param>
        /// <returns></returns>
        public ExcelChart Add(eChartType chartType)
        {
            if (_topChart.PivotTableSource != null)
            {
                throw (new InvalidOperationException("Can not add other charttypes to a pivot chart"));
            }
            else if (ExcelChart.IsType3D(chartType) || _list[0].IsType3D())
            {
                throw(new InvalidOperationException("3D charts can not be combined with other charttypes"));
            }

            var prependingChartNode = _list[_list.Count - 1].TopNode;
            ExcelChart chart = ExcelChart.GetNewChart(_topChart.WorkSheet.Drawings, _topChart.TopNode, chartType, _topChart, null);

            _list.Add(chart);
            return chart;
        }
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }
        IEnumerator<ExcelChart> IEnumerable<ExcelChart>.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        /// <summary>
        /// Returns a chart at the specific position.  
        /// </summary>
        /// <param name="PositionID">The position of the chart. 0-base</param>
        /// <returns></returns>
        public ExcelChart this[int PositionID]
        {
            get
            {
                return (_list[PositionID]);
            }
        }


}
}
