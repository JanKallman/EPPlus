using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlusSamples
{
    class Sample_ErrorBars
    {
        public static void RunSample_ErrorBars()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Error Bars");

                var txt = "Sample,X,Y,X +error,X -error,Y +error,Y -error\r\n" +
                "A,5,3,0.5,0.9,0.3,0.7\r\n" +
                "B,6,4,0.6,0.8,0.4,0.6\r\n" +
                "C,7,5,0.7,0.7,0.5,0.5\r\n" +
                "D,8,6,0.8,0.6,0.6,0.4\r\n" +
                "E,9,7,0.9,0.5,0.7,0.3\r\n";

                ws.Cells["A1"].LoadFromText(txt);

                #region Add Column chart
                {
                    var columnChart = (ExcelBarChart)ws.Drawings.AddChart("ColumnChart1", eChartType.ColumnClustered);
                    var columnSeries = (ExcelBarChartSerie)columnChart.Series.Add(ExcelCellBase.GetAddress(2, 2, 6, 2), ExcelCellBase.GetAddress(2, 1, 6, 1));
                    columnChart.Style = eChartStyle.Style2;
                    columnChart.SetPosition(8, 0, 0, 0);

                    columnSeries.ErrorBar.Type = eErrorBarType.Both;
                    columnSeries.ErrorBar.ValueType = eErrorBarValueType.CustomErrorBars;
                    columnSeries.ErrorBar.NoEndCap = false;

                    columnSeries.ErrorBar.PlusAddress = "D2:D6";
                    columnSeries.ErrorBar.MinusAddress = "E2:E6";

                    columnSeries.ErrorBar.Line.Fill.Color = System.Drawing.Color.Red;
                }
                #endregion

                #region Add Bar chart
                {
                    var barChart = (ExcelBarChart)ws.Drawings.AddChart("BarChart1", eChartType.BarClustered);
                    var barSeries = (ExcelBarChartSerie)barChart.Series.Add(ExcelCellBase.GetAddress(2, 2, 6, 2), ExcelCellBase.GetAddress(2, 1, 6, 1));
                    barChart.Style = eChartStyle.Style2;
                    barChart.SetPosition(19, 0, 0, 0);

                    barSeries.ErrorBar.Type = eErrorBarType.Plus;
                    barSeries.ErrorBar.ValueType = eErrorBarValueType.Percentage;
                    barSeries.ErrorBar.Value = 10;
                    barSeries.ErrorBar.NoEndCap = false;
                }

                #endregion

                package.SaveAs(Utils.GetFileInfo("Sample_ErrorBars.xlsx"));
            }
        }
    }
}
