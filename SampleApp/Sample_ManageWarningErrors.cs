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
 * See https://github.com/JanKallman/EPPlus for details.
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
    class Sample_ManageWarningErrors
    {
        /// <summary>
        /// Sample ManageWarningErrors - Shows how to enable/disable cells warning errors, like "Number Stored as Text" or "Two Digit Text Year".
        /// Useful when you want to display "numbers" like excel detects as text (e.g. numbers leading by zeros)
        /// The workbook contains worksheets with / without warnings
        /// </summary>
        public static string RunSample_ManageWarningErrors()
        {
            using (var package = new ExcelPackage())
            {
                // Add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = CreateSampleBaseWorkSheet(package, "Warnings-ON");

                ExcelWorksheet worksheet2 = CreateSampleBaseWorkSheet(package, "Warnings-OFF");
                // Set the cell range to ignore errors on to the entire worksheet
                worksheet2.IgnoredError.Range = worksheet2.Dimension.Address;

                // Do not display the warning 'number stored as text'
                worksheet2.IgnoredError.NumberStoredAsText = true;

                ExcelWorksheet worksheet3 = CreateSampleBaseWorkSheet(package, "Warnings-Selective");
                // Set the cell range to ignore errors only on to column B 
                worksheet3.IgnoredError.Range = "B:B";

                // Do not display the warning 'number stored as text'
                worksheet3.IgnoredError.NumberStoredAsText = true;

                // set some document properties
                package.Workbook.Properties.Title = "Check Warnings Sample";
                package.Workbook.Properties.Author = "romcode";
                package.Workbook.Properties.Comments = "This sample demonstrates how to use the managing errors feature using EPPlus";

                var xlFile = Utils.GetFileInfo("sample_warning_errors.xlsx");
                // save our new workbook in the output directory and we are done!
                package.SaveAs(xlFile);
                return xlFile.FullName;
            }
        }

        private static ExcelWorksheet CreateSampleBaseWorkSheet(ExcelPackage package,string name)
        {
            // Create the worksheet
            Console.WriteLine("Creating worksheet {0}", name);

            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(name);

            // Add the headers
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Tag Code";
            worksheet.Cells[1, 3].Value = "Product";

            //Add some items...
            worksheet.Cells["A2"].Value = "01";
            worksheet.Cells["B2"].Value = "00001";
            worksheet.Cells["C2"].Value = "Nails";

            worksheet.Cells["A3"].Value = "02";
            worksheet.Cells["B3"].Value = "00002";
            worksheet.Cells["C3"].Value = "Books";

            worksheet.Cells["A4"].Value = "03";
            worksheet.Cells["B4"].Value = "00003";
            worksheet.Cells["C4"].Value = "Fruits";

            return worksheet;
        }
    }
}
