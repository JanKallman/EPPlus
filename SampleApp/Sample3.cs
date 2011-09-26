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

/*
 * Sample code demonstrating how to generate Excel spreadsheets on the server using 
 * Office Open XML and the ExcelPackage wrapper classes.
 * 
 * ExcelPackage provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/ExcelPackage for details.
 * 
 * Sample 3: Creates a workbook based on a template and populates using the database data.
 * 
 * Copyright 2007 © Dr John Tunnicliffe 
 * mailto:dr.john.tunnicliffe@btinternet.com
 * All rights reserved.
 * 
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 */
using System;
using System.IO;
using System.Xml;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Drawing;
using OfficeOpenXml.Style;

namespace EPPlusSamples
{
	class Sample3
	{
		/// <summary>
		/// Sample 3 - creates a workbook and populates using data from the AdventureWorks database
		/// This sample requires the AdventureWorks database.  
        /// This one is from the orginal Excelpackage sample project, but without the template
		/// </summary>
		/// <param name="outputDir">The output directory</param>
		/// <param name="templateDir">The location of the sample template</param>
		/// <param name="connectionString">The connection string to your copy of the AdventureWorks database</param>
		public static string RunSample3(DirectoryInfo outputDir, string connectionString)
		{
			
            string file = outputDir.FullName + @"\sample3.xlsx";
            if (File.Exists(file)) File.Delete(file);
            FileInfo newFile = new FileInfo(outputDir.FullName + @"\sample3.xlsx");

			// ok, we can run the real code of the sample now
			using (ExcelPackage xlPackage = new ExcelPackage(newFile))
			{
				// uncomment this line if you want the XML written out to the outputDir
				//xlPackage.DebugMode = true; 

				// get handle to the existing worksheet
				ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.Add("Sales");
                var namedStyle = xlPackage.Workbook.Styles.CreateNamedStyle("HyperLink");   //This one is language dependent
                namedStyle.Style.Font.UnderLine = true;
                namedStyle.Style.Font.Color.SetColor(Color.Blue);
                if (worksheet != null)
				{
					const int startRow = 5;
					int row = startRow;
                    //Create Headers and format them 
                    worksheet.Cells["A1"].Value = "AdventureWorks Inc.";
                    using (ExcelRange r = worksheet.Cells["A1:G1"])
                    {
                        r.Merge = true;
                        r.Style.Font.SetFromFont(new Font("Britannic Bold", 22, FontStyle.Italic));
                        r.Style.Font.Color.SetColor(Color.White);
                        r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                        r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23,55,93));
                    }
                    worksheet.Cells["A2"].Value = "Year-End Sales Report";
                    using (ExcelRange r = worksheet.Cells["A2:G2"])
                    {
                        r.Merge = true;
                        r.Style.Font.SetFromFont(new Font("Britannic Bold", 18, FontStyle.Italic));
                        r.Style.Font.Color.SetColor(Color.Black);
                        r.Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                        r.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184,204,228));
                    }

                    worksheet.Cells["A4"].Value = "Name";
                    worksheet.Cells["B4"].Value = "Job Title";
                    worksheet.Cells["C4"].Value = "Region";
                    worksheet.Cells["D4"].Value = "Monthly Quota";
                    worksheet.Cells["E4"].Value = "Quota YTD";
                    worksheet.Cells["F4"].Value = "Sales YTD";
                    worksheet.Cells["G4"].Value = "Quota %";
                    worksheet.Cells["A4:G4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells["A4:G4"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                    worksheet.Cells["A4:G4"].Style.Font.Bold = true;


                    // lets connect to the AdventureWorks sample database for some data
					using (SqlConnection sqlConn = new SqlConnection(connectionString))
					{
						sqlConn.Open();
						using (SqlCommand sqlCmd = new SqlCommand("select LastName + ', ' + FirstName AS [Name], EmailAddress, JobTitle, CountryRegionName, ISNULL(SalesQuota,0) AS SalesQuota, ISNULL(SalesQuota,0)*12 AS YearlyQuota, SalesYTD from Sales.vSalesPerson ORDER BY SalesYTD desc", sqlConn))
						{
							using (SqlDataReader sqlReader = sqlCmd.ExecuteReader())
							{
								// get the data and fill rows 5 onwards
								while (sqlReader.Read())
								{
									int col = 1;
									// our query has the columns in the right order, so simply
									// iterate through the columns
									for (int i = 0; i < sqlReader.FieldCount; i++)
									{
										// use the email address as a hyperlink for column 1
										if (sqlReader.GetName(i) == "EmailAddress")
										{
											// insert the email address as a hyperlink for the name
											string hyperlink = "mailto:" + sqlReader.GetValue(i).ToString();
											worksheet.Cells[row, 1].Hyperlink = new Uri(hyperlink, UriKind.Absolute);
										}
										else
										{
											// do not bother filling cell with blank data (also useful if we have a formula in a cell)
											if (sqlReader.GetValue(i) != null)
												worksheet.Cells[row, col].Value = sqlReader.GetValue(i);
											col++;
										}
									}
									row++;
								}
								sqlReader.Close();

                                worksheet.Cells[startRow, 1, row - 1, 1].StyleName = "HyperLink";
                                worksheet.Cells[startRow, 4, row - 1, 6].Style.Numberformat.Format = "[$$-409]#,##0";
                                worksheet.Cells[startRow, 7, row - 1, 7].Style.Numberformat.Format = "0%";

                                worksheet.Cells[startRow, 7, row - 1, 7].FormulaR1C1 = "=IF(RC[-2]=0,0,RC[-1]/RC[-2])";

                                //Set column width
                                worksheet.Column(1).Width = 25;
                                worksheet.Column(2).Width = 28;
                                worksheet.Column(3).Width = 18;
                                worksheet.Column(4).Width = 12;
                                worksheet.Column(5).Width = 10;
                                worksheet.Column(6).Width = 10;
                                worksheet.Column(7).Width = 12;
                            }
						}
						sqlConn.Close();
					}

                    // lets set the header text 
					worksheet.HeaderFooter.OddHeader.CenteredText = "AdventureWorks Inc. Sales Report";
					// add the page number to the footer plus the total number of pages
					worksheet.HeaderFooter.OddFooter.RightAlignedText =
						string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
					// add the sheet name to the footer
					worksheet.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
					// add the file path to the footer
					worksheet.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FilePath + ExcelHeaderFooter.FileName;
				}
				// we had better add some document properties to the spreadsheet 

				// set some core property values
				xlPackage.Workbook.Properties.Title = "Sample 3";
				xlPackage.Workbook.Properties.Author = "John Tunnicliffe";
				xlPackage.Workbook.Properties.Subject = "ExcelPackage Samples";
				xlPackage.Workbook.Properties.Keywords = "Office Open XML";
				xlPackage.Workbook.Properties.Category = "ExcelPackage Samples";
				xlPackage.Workbook.Properties.Comments = "This sample demonstrates how to create an Excel 2007 file from scratch using the Packaging API and Office Open XML";

				// set some extended property values
				xlPackage.Workbook.Properties.Company = "AdventureWorks Inc.";
                xlPackage.Workbook.Properties.HyperlinkBase = new Uri("http://www.codeplex.com/MSFTDBProdSamples");

				// set some custom property values
				xlPackage.Workbook.Properties.SetCustomPropertyValue("Checked by", "John Tunnicliffe");
				xlPackage.Workbook.Properties.SetCustomPropertyValue("EmployeeID", "1147");
				xlPackage.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "ExcelPackage");
				
				// save the new spreadsheet
				xlPackage.Save();
			}

			// if you want to take a look at the XML created in the package, simply uncomment the following lines
			// These copy the output file and give it a zip extension so you can open it and take a look!
			//FileInfo zipFile = new FileInfo(outputDir.FullName + @"\sample3.zip");
			//if (zipFile.Exists) zipFile.Delete();
			//newFile.CopyTo(zipFile.FullName);

			return newFile.FullName;
		}
	}
}
