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
 * Jan Källman		Added		07-JAN-2010
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;    
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing;
using System.Drawing;

namespace EPPlusSamples
{
    class Sample5
    {
        /// <summary>
        /// Sample 5 - open Sample 1 and add 2 new rows and a Piechart
        /// </summary>
        public static string RunSample5(DirectoryInfo outputDir)
        {
            FileInfo templateFile = new FileInfo(outputDir.FullName + @"\sample1.xlsx");
            FileInfo newFile = new FileInfo(outputDir.FullName + @"\sample5.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(outputDir.FullName + @"\sample5.xlsx");
            }
            using (ExcelPackage package = new ExcelPackage(newFile, templateFile))
            {
                //Open worksheet 1
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                worksheet.InsertRow(5, 2);

                worksheet.Cells["A5"].Value = "12010";
                worksheet.Cells["B5"].Value = "Drill";
                worksheet.Cells["C5"].Value = 20;
                worksheet.Cells["D5"].Value = 8;

                worksheet.Cells["A6"].Value = "12011";
                worksheet.Cells["B6"].Value = "Crowbar";
                worksheet.Cells["C6"].Value = 7;
                worksheet.Cells["D6"].Value = 23.48;

                worksheet.Cells["E2:E6"].FormulaR1C1 = "RC[-2]*RC[-1]";                

                var name = worksheet.Names.Add("SubTotalName", worksheet.Cells["C7:E7"]);
                name.Style.Font.Italic = true;
                name.Formula = "SUBTOTAL(9,C2:C6)";

                //Format the new rows
                worksheet.Cells["C5:C6"].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["D5:E6"].Style.Numberformat.Format = "#,##0.00";

                var chart = (worksheet.Drawings.AddChart("PieChart", eChartType.Pie3D) as ExcelPieChart);

                chart.Title.Text = "Total";
                //From row 1 colum 5 with five pixels offset
                chart.SetPosition(0, 0, 5, 5);
                chart.SetSize(600, 300);

                ExcelAddress valueAddress = new ExcelAddress(2, 5, 6, 5);
                var ser = (chart.Series.Add(valueAddress.Address, "B2:B6") as ExcelPieChartSerie);
                chart.DataLabel.ShowCategory = true;
                chart.DataLabel.ShowPercent = true;

                chart.Legend.Border.LineStyle = eLineStyle.Solid;
                chart.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chart.Legend.Border.Fill.Color = Color.DarkBlue;

                //Switch the PageLayoutView back to normal
                worksheet.View.PageLayoutView = false;
                // save our new workbook and we are done!
                package.Save();
            }

            return newFile.FullName;
        }
    }
}
