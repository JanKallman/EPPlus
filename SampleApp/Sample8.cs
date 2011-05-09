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
using System.Linq;
namespace EPPlusSamples
{
    public static class LinqSample
    {
        /// <summary>
        /// This sample shows how to use Linq with the Cells collection
        /// </summary>
        /// <param name="outputDir">The path where sample7.xlsx is</param>
        public static void RunLinqSample(DirectoryInfo outputDir)
        {
	        Console.WriteLine("Now open sample 7 again and perform some Linq queries...");
		    Console.WriteLine();

			FileInfo existingFile = new FileInfo(outputDir.FullName + @"\sample7.xlsx");
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets[1];
                
                //Select all cells in column d between 9990 and 10000
                var query1= (from cell in sheet.Cells["d:d"] where cell.Value is double && (double)cell.Value >= 9990 && (double)cell.Value <= 10000 select cell);

                Console.WriteLine("Print all cells with value between 9990 and 10000 in column D ...");
                Console.WriteLine();

                int count = 0;
                foreach (var cell in query1)
                {
                    Console.WriteLine("Cell {0} has value {1:N0}", cell.Address, cell.Value);
                    count++;
                }

                Console.WriteLine("{0} cells found ...",count);
                Console.WriteLine();

                //Select all bold cells
                Console.WriteLine("Now get all bold cells from the entire sheet...");
                var query2 = (from cell in sheet.Cells[sheet.Dimension.Address] where cell.Style.Font.Bold select cell);
                //If you have a clue where the data is, specify a smaller range in the cells indexer to get better performance (for example "1:1,65536:65536" here)
                count = 0;
                foreach (var cell in query2)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        Console.WriteLine("Cell {0} is bold and has a formula of {1:N0}", cell.Address, cell.Formula);
                    }
                    else
                    {
                        Console.WriteLine("Cell {0} is bold and has a value of {1:N0}", cell.Address, cell.Value);
                    }
                    count++;
                }

                //Here we use more than one column in the where clause. We start by searching column D, then use the Offset method to check the value of column C.
                var query3 = (from cell in sheet.Cells["d:d"]
                              where cell.Value is double && 
                                    (double)cell.Value >= 9500 && (double)cell.Value <= 10000 && 
                                    cell.Offset(0, -1).GetValue<DateTime>().Year == DateTime.Today.Year+1 
                              select cell);

                Console.WriteLine();
                Console.WriteLine("Print all cells with a value between 9500 and 10000 in column D and the year of Column C is {0} ...", DateTime.Today.Year + 1);
                Console.WriteLine();    

                count = 0;
                foreach (var cell in query3)    //The cells returned here will all be in column D, since that is the address in the indexer. Use the Offset method to print any other cells from the same row.
                {
                    Console.WriteLine("Cell {0} has value {1:N0} Date is {2:d}", cell.Address, cell.Value, cell.Offset(0, -1).GetValue<DateTime>());
                    count++;
                }
            }
        }
    }
}
