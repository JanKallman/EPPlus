using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{

    [TestClass]
    public class ExcelChartDataTableTest
    {
        public TestContext TestContext { get; set; }


        /// <summary>
        /// Basic test to check output with excel. need enhanced to be stand alone checking
        /// </summary>
        [TestMethod]
        [Ignore]
        public void DataTableFile()
        {
            string outfile = Path.Combine(TestContext.TestResultsDirectory, "DataTableFile.xlsx");
            var fileinfo = new FileInfo(outfile);
            using (ExcelPackage pkg = new ExcelPackage(fileinfo))
            {
                // Add worksheet with sample data
                var worksheet = pkg.Workbook.Worksheets.Add("TestData");
                worksheet.Cells["A1"].Value = "Data";
                for (int x = 1; x < 12; ++x)
                {
                    worksheet.Cells[x+1, 1].Value = (double)x / 3.0;
                }

                // Add chart from sample data
                var chartsheet = pkg.Workbook.Worksheets.AddChart("TestChart", eChartType.Line);
                var chart = chartsheet.Chart as ExcelLineChart;
                chart.Series.Add(worksheet.Cells["A1:A12"], worksheet.Cells["A1:A12"]).Header = "Data Test";

                // as per epplus style, data table does not init until used. so call something
                chart.PlotArea.DataTable.ShowKeys = true;

                pkg.Save();
            }
        }
    }
}
