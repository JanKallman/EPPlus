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
 * EPPlus is a fork of the ExcelPackage project
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

namespace ExcelPackageSamples
{
    class Sample5
    {
        /// <summary>
        /// Sample 5 - open Sample 1 and add a Piechart
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
            using (ExcelPackage xlPackage = new ExcelPackage(newFile, templateFile))
            {
                // this will cause the assembly to output the raw XML files in the outputDir
                // for debug purposes.  You will see to sub-folders called 'xl' and 'docProps'.
                xlPackage.DebugMode = true;

                //Open worksheet 1
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[1];
                var chart = (worksheet.Drawings.AddChart("PieChart", eChartType.Pie3D) as ExcelPieChart);

                chart.Title.Text = "Total";
                chart.SetPosition(0, 0, 2, 5);
                chart.SetSize(600, 300);

                string valueAddress = ExcelRange.GetAddress(2, 2, 4, 2);
                var ser = (chart.Series.Add(valueAddress, "A2:A4") as ExcelPieChartSerie);
                chart.DataLabel.ShowCategory = true;
                chart.DataLabel.ShowPercent = true;

                chart.Legend.Border.LineStyle = eLineStyle.Solid;
                chart.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chart.Legend.Border.Fill.Color = Color.DarkBlue;
                
                worksheet.View.PageLayoutView = false;
                // save our new workbook and we are done!
                xlPackage.Save();
            }

            // if you want to take a look at the XML created in the package, simply uncomment the following lines
            // These copy the output file and give it a zip extension so you can open it and take a look!
            //FileInfo zipFile = new FileInfo(outputDir.FullName + @"\sample1.zip");
            //if (zipFile.Exists) zipFile.Delete();
            //newFile.CopyTo(zipFile.FullName);
            return newFile.FullName;
        }
    }
}
