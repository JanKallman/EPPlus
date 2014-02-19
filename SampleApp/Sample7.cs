using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
namespace EPPlusSamples
{
    class Sample7
    {
        /// <summary>
        /// This sample load a number of rows, style them and insert a row at the top.
        /// A password is set to protect locked cells. Column 3 & 4 will be editable, the rest will be locked.
        /// </summary>
        /// <param name="outputDir"></param>
        /// <param name="Rows"></param>
        public static string RunSample7(DirectoryInfo outputDir, int Rows)
        {
            FileInfo newFile = new FileInfo(outputDir.FullName + @"\sample7.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(outputDir.FullName + @"\sample7.xlsx");
            }

            using (ExcelPackage package = new ExcelPackage())
            {
                Console.WriteLine("{0:HH.mm.ss}\tStarting...", DateTime.Now);

                //Load the sheet with one string column, one date column and a few random numbers.
                var ws = package.Workbook.Worksheets.Add("Performance Test");

                //Format all cells
                ExcelRange cols = ws.Cells["A:XFD"];
                cols.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cols.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                var rnd = new Random();                
                for (int row = 1; row <= Rows; row++)
                {
                    ws.SetValue(row, 1, row);                               //The SetValue method is a little bit faster than using the Value property
                    ws.SetValue(row, 2, string.Format("Row {0}", row));
                    ws.SetValue(row, 3, DateTime.Today.AddDays(row));
                    ws.SetValue(row, 4, rnd.NextDouble() * 10000);
                    if (row % 10000 == 0)
                    {
                        Console.WriteLine("{0:HH.mm.ss}\tWriting row {1}...", DateTime.Now, row);
                    }
                }
                ws.Cells[1, 5, Rows, 5].FormulaR1C1 = "RC[-4]+RC[-1]";

                //Add a sum at the end
                ws.Cells[Rows + 1, 5].Formula = string.Format("Sum({0})", new ExcelAddress(1, 5, Rows, 5).Address);
                ws.Cells[Rows + 1, 5].Style.Font.Bold = true;
                ws.Cells[Rows + 1, 5].Style.Numberformat.Format = "#,##0.00";

                Console.WriteLine("{0:HH.mm.ss}\tWriting row {1}...", DateTime.Now, Rows);
                Console.WriteLine("{0:HH.mm.ss}\tFormatting...", DateTime.Now);
                //Format the date and numeric columns
                ws.Cells[1, 1, Rows, 1].Style.Numberformat.Format = "#,##0";
                ws.Cells[1, 3, Rows, 3].Style.Numberformat.Format = "YYYY-MM-DD";
                ws.Cells[1, 4, Rows, 5].Style.Numberformat.Format = "#,##0.00";

                Console.WriteLine("{0:HH.mm.ss}\tInsert a row at the top...", DateTime.Now);
                //Insert a row at the top. Note that the formula-addresses are shifted down
                ws.InsertRow(1, 1);

                //Write the headers and style them
                ws.Cells["A1"].Value = "Index";
                ws.Cells["B1"].Value = "Text";
                ws.Cells["C1"].Value = "Date";
                ws.Cells["D1"].Value = "Number";
                ws.Cells["E1"].Value = "Formula";
                ws.View.FreezePanes(2, 1);

                using (var rng = ws.Cells["A1:E1"])
                {
                    rng.Style.Font.Bold = true;
                    rng.Style.Font.Color.SetColor(Color.White);
                    rng.Style.WrapText = true;
                    rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                }

                //Calculate
                Console.WriteLine("{0:HH.mm.ss}\tCalculate formulas...", DateTime.Now);
                ws.Calculate();

                Console.WriteLine("{0:HH.mm.ss}\tAutofit columns and lock and format cells...", DateTime.Now);
                ws.Cells[Rows - 100, 1, Rows, 5].AutoFitColumns(5);   //Auto fit using the last 100 rows with minimum width 5
                ws.Column(5).Width = 15;                            //We need to set the width for column F manually since the end sum formula is the widest cell in the column (EPPlus don't calculate any forumlas, so no output text is avalible). 

                //Now we set the sheetprotection and a password.
                ws.Cells[2, 3, Rows + 1, 4].Style.Locked = false;
                ws.Cells[2, 3, Rows + 1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[2, 3, Rows + 1, 4].Style.Fill.BackgroundColor.SetColor(Color.White);
                ws.Cells[1, 5, Rows + 2, 5].Style.Hidden = true;    //Hide the formula
                
                ws.Protection.SetPassword("EPPlus");

                ws.Select("C2");
                Console.WriteLine("{0:HH.mm.ss}\tSaving...", DateTime.Now);
                package.Compression = CompressionLevel.BestSpeed;
                package.SaveAs(newFile);
            }
            Console.WriteLine("{0:HH.mm.ss}\tDone!!", DateTime.Now);
            return newFile.FullName;
        }
    }
}
