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
 * Jan Källman		Added		21 Mar 2010
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
namespace EPPlusSamples
{
	class Sample1
	{
		/// <summary>
		/// Sample 1 - simply creates a new workbook from scratch.
		/// The workbook contains one worksheet with a simple invertory list
		/// </summary>
		public static string RunSample1(DirectoryInfo outputDir)
		{
			FileInfo newFile = new FileInfo(outputDir.FullName + @"\sample1.xlsx");
			if (newFile.Exists)
			{
				newFile.Delete();  // ensures we create a new workbook
				newFile = new FileInfo(outputDir.FullName + @"\sample1.xlsx");
			}
			using (ExcelPackage package = new ExcelPackage(newFile))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Inventory");
                //Add the headers
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Product";
                worksheet.Cells[1, 3].Value = "Quantity";
                worksheet.Cells[1, 4].Value = "Price";
                worksheet.Cells[1, 5].Value = "Value";

                //Add some items...
                worksheet.Cells["A2"].Value = 12001;
                worksheet.Cells["B2"].Value = "Nails";
                worksheet.Cells["C2"].Value = 37;
                worksheet.Cells["D2"].Value = 3.99;

                worksheet.Cells["A3"].Value = 12002;
                worksheet.Cells["B3"].Value = "Hammer";
                worksheet.Cells["C3"].Value = 5;
                worksheet.Cells["D3"].Value = 12.10;

                worksheet.Cells["A4"].Value = 12003;
                worksheet.Cells["B4"].Value = "Saw";
                worksheet.Cells["C4"].Value = 12;
                worksheet.Cells["D4"].Value = 15.37;

                //Add a formula for the value-column
                worksheet.Cells["E2:E4"].Formula = "C2*D2";

                //Ok now format the values;
                using (var range = worksheet.Cells[1, 1, 1, 5]) 
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                    range.Style.Font.Color.SetColor(Color.White);
                }

                worksheet.Cells["A5:E5"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["A5:E5"].Style.Font.Bold = true;

                worksheet.Cells[5, 3, 5, 5].Formula = string.Format("SUBTOTAL(9,{0})", new ExcelAddress(2,3,4,3).Address);
                worksheet.Cells["C2:C5"].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["D2:E5"].Style.Numberformat.Format = "#,##0.00";
                
                //Create an autofilter for the range
                worksheet.Cells["A1:E4"].AutoFilter = true;

                worksheet.Cells["A2:A4"].Style.Numberformat.Format = "@";   //Format as text

                //There is actually no need to calculate, Excel will do it for you, but in some cases it might be useful. 
                //For example if you link to this workbook from another workbook or you will open the workbook in a program that hasn't a calculation engine or 
                //you want to use the result of a formula in your program.
                worksheet.Calculate(); 

                worksheet.Cells.AutoFitColumns(0);  //Autofit columns for all cells

                // lets set the header text 
                worksheet.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\" Inventory";
                // add the page number to the footer plus the total number of pages
                worksheet.HeaderFooter.OddFooter.RightAlignedText =
                    string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                // add the sheet name to the footer
                worksheet.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                // add the file path to the footer
                worksheet.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FilePath + ExcelHeaderFooter.FileName;

                worksheet.PrinterSettings.RepeatRows = worksheet.Cells["1:2"];
                worksheet.PrinterSettings.RepeatColumns = worksheet.Cells["A:G"];

                // Change the sheet view to show it in page layout mode
                worksheet.View.PageLayoutView = true;

                // set some document properties
                package.Workbook.Properties.Title = "Invertory";
                package.Workbook.Properties.Author = "Jan Källman";
                package.Workbook.Properties.Comments = "This sample demonstrates how to create an Excel 2007 workbook using EPPlus";

                // set some extended property values
                package.Workbook.Properties.Company = "AdventureWorks Inc.";

                // set some custom property values
                package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Jan Källman");
                package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus");
                // save our new workbook and we are done!
                package.Save();

            }

            return newFile.FullName;
		}
	}
}
