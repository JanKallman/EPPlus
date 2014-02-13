/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 *
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
 *******************************************************************************
 * Jan Källman		Added		28 Oct 2010
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Table;
using OfficeOpenXml.Drawing.Chart;
using System.Globalization;
namespace EPPlusSamples
{
    /// <summary>
    /// This sample shows how to load CSV files using the LoadFromText method, how to use tables and
    /// how to use charts with more than one charttype and secondary axis
    /// </summary>
    public static class Sample9
    {
        /// <summary>
        /// Loads two CSV files into tables and adds a chart to each sheet.
        /// </summary>
        /// <param name="outputDir"></param>
        /// <returns></returns>
        public static string RunSample9(DirectoryInfo outputDir)
        {
            FileInfo newFile = new FileInfo(outputDir.FullName + @"\sample9.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(outputDir.FullName + @"\sample9.xlsx");
            }
            
            using (ExcelPackage package = new ExcelPackage())
            {
                LoadFile1(package);
                LoadFile2(package);

                package.SaveAs(newFile);
            }
            return newFile.FullName;
        }
        private static void LoadFile1(ExcelPackage package)
        {
            //Create the Worksheet
            var sheet = package.Workbook.Worksheets.Add("Csv1");

            //Create the format object to describe the text file
            var format = new ExcelTextFormat();
            format.TextQualifier = '"';
            format.SkipLinesBeginning = 2;
            format.SkipLinesEnd = 1;

            //Now read the file into the sheet. Start from cell A1. Create a table with style 27. First row contains the header.
            Console.WriteLine("Load the text file...");
            var range = sheet.Cells["A1"].LoadFromText(new FileInfo("..\\..\\csv\\Sample9-1.txt"), format, TableStyles.Medium27, true);

            Console.WriteLine("Format the table...");
            //Tables don't support custom styling at this stage(you can of course format the cells), but we can create a Namedstyle for a column...
            var dateStyle = package.Workbook.Styles.CreateNamedStyle("TableDate");
            dateStyle.Style.Numberformat.Format = "YYYY-MM";

            var numStyle = package.Workbook.Styles.CreateNamedStyle("TableNumber");
            numStyle.Style.Numberformat.Format = "#,##0.0";

            //Now format the table...
            var tbl = sheet.Tables[0];
            tbl.ShowTotal = true;
            tbl.Columns[0].TotalsRowLabel = "Total";
            tbl.Columns[0].DataCellStyleName = "TableDate";
            tbl.Columns[1].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[1].DataCellStyleName = "TableNumber";
            tbl.Columns[2].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[2].DataCellStyleName = "TableNumber";
            tbl.Columns[3].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[3].DataCellStyleName = "TableNumber";
            tbl.Columns[4].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[4].DataCellStyleName = "TableNumber";
            tbl.Columns[5].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[5].DataCellStyleName = "TableNumber";
            tbl.Columns[6].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[6].DataCellStyleName = "TableNumber";
            
            Console.WriteLine("Create the chart...");
            //Now add a stacked areachart...
            var chart = sheet.Drawings.AddChart("chart1", eChartType.AreaStacked);
            chart.SetPosition(0, 630);
            chart.SetSize(800, 600);

            //Create one series for each column...
            for (int col = 1; col < 7; col++)
            {
                var ser = chart.Series.Add(range.Offset(1, col, range.End.Row - 1, 1), range.Offset(1, 0, range.End.Row - 1, 1));
                ser.HeaderAddress = range.Offset(0, col, 1, 1);
            }
            
            //Set the style to 27.
            chart.Style = eChartStyle.Style27;

            sheet.View.ShowGridLines = false;
            sheet.Calculate();
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
        }

        private static void LoadFile2(ExcelPackage package)
        {
            //Create the Worksheet
            var sheet = package.Workbook.Worksheets.Add("Csv2");

            //Create the format object to describe the text file
            var format = new ExcelTextFormat();
            format.Delimiter='\t'; //Tab
            format.SkipLinesBeginning = 1;
            CultureInfo ci = new CultureInfo("sv-SE");          //Use your choice of Culture
            ci.NumberFormat.NumberDecimalSeparator = ",";       //Decimal is comma
            format.Culture = ci;

            //Now read the file into the sheet.
            Console.WriteLine("Load the text file...");
            var range = sheet.Cells["A1"].LoadFromText(new FileInfo("..\\..\\csv\\Sample9-2.txt"), format);

            //Add a formula
            range.Offset(1, range.End.Column, range.End.Row - range.Start.Row, 1).FormulaR1C1 = "RC[-1]-RC[-2]";

            //Add a table...
            var tbl = sheet.Tables.Add(range.Offset(0,0,range.End.Row-range.Start.Row+1, range.End.Column-range.Start.Column+2),"Table");
            tbl.ShowTotal = true;
            tbl.Columns[0].TotalsRowLabel = "Total";
            tbl.Columns[1].TotalsRowFormula = "COUNT(3,Table[Product])";    //Add a custom formula
            tbl.Columns[2].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[3].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[4].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[5].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[5].Name = "Profit";
            tbl.TableStyle = TableStyles.Medium10;

            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();

            //Add a chart with two charttypes (Column and Line) and a secondary axis...
            var chart = sheet.Drawings.AddChart("chart2", eChartType.ColumnStacked);
            chart.SetPosition(0, 540);
            chart.SetSize(800, 600);
        
            var serie1= chart.Series.Add(range.Offset(1, 3, range.End.Row - 1, 1), range.Offset(1, 1, range.End.Row - 1, 1));
            serie1.Header = "Purchase Price";
            var serie2 = chart.Series.Add(range.Offset(1, 5, range.End.Row - 1, 1), range.Offset(1, 1, range.End.Row - 1, 1));
            serie2.Header = "Profit";

            //Add a Line series
            var chartType2 = chart.PlotArea.ChartTypes.Add(eChartType.LineStacked);
            chartType2.UseSecondaryAxis = true;
            var serie3 = chartType2.Series.Add(range.Offset(1, 2, range.End.Row - 1, 1), range.Offset(1, 0, range.End.Row - 1, 1));
            serie3.Header = "Items in stock";

            //By default the secondary XAxis is not visible, but we want to show it...
            chartType2.XAxis.Deleted = false;
            chartType2.XAxis.TickLabelPosition = eTickLabelPosition.High;
            
            //Set the max value for the Y axis...
            chartType2.YAxis.MaxValue = 50;

            chart.Style = eChartStyle.Style26;
            sheet.View.ShowGridLines = false;
            sheet.Calculate();
        }
    }
}
