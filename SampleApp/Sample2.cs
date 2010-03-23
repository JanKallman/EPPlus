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
 * Jan Källman		Added		10-SEP-2009
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;

namespace EPPlusSamples
{
	/// <summary>
	/// Simply opens an existing file and reads some values and properties
	/// </summary>
	class Sample2
	{
		public static void RunSample2(string FilePath)
		{
			Console.WriteLine("Reading column 2 of {0}", FilePath);
			Console.WriteLine();

			FileInfo existingFile = new FileInfo(FilePath);
			using (ExcelPackage package = new ExcelPackage(existingFile))
			{
				// get the first worksheet in the workbook
				ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                int col = 2; //The item description
				// output the data in column 2
				for (int row = 2; row < 5; row++)
					Console.WriteLine("\tCell({0},{1}).Value={2}", row, col, worksheet.Cells[row, col].Value);

				// output the formula in row 5
				Console.WriteLine("\tCell({0},{1}).Formula={2}", 3, 5, worksheet.Cells[3, 5].Formula);                
                Console.WriteLine("\tCell({0},{1}).FormulaR1C1={2}", 3, 5, worksheet.Cells[3, 5].FormulaR1C1);

                // output the formula in row 5
                Console.WriteLine("\tCell({0},{1}).Formula={2}", 5, 3, worksheet.Cells[5, 3].Formula);
                Console.WriteLine("\tCell({0},{1}).FormulaR1C1={2}", 5, 3, worksheet.Cells[5, 3].FormulaR1C1);

            } // the using statement automatically calls Dispose() which closes the package.

			Console.WriteLine();
			Console.WriteLine("Sample 2 complete");
			Console.WriteLine();
		}
	}
}
