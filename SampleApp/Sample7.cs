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
        public static void RunSample7(DirectoryInfo outputDir, int Rows)
        {
            FileInfo newFile = new FileInfo(outputDir.FullName + @"\sample7.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(outputDir.FullName + @"\sample7.xlsx");
            }

            ExcelPackage package = new ExcelPackage();
            Console.WriteLine("{0:HH.mm.ss}\tStarting...", DateTime.Now);

            //Load the sheet with one string column, one date column and a few random numbers.
            var ws = package.Workbook.Worksheets.Add("Performance Test");

            //This is a trick to format all cells in a sheet.
            //Note that this only work when no columns are changed (added)
            ExcelColumn cols = ws.Column(1);
            cols.ColumnMax = ExcelPackage.MaxColumns;
            cols.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cols.Style.Fill.BackgroundColor.SetColor(Color.LightGray);

            var rnd = new Random();
            for (int row = 1; row <= Rows; row++)
            {
                ws.Cells[row, 1].Value = row;
                ws.Cells[row, 2].Value = string.Format("Row {0}", row);
                ws.Cells[row, 3].Value = DateTime.Today.AddDays(row);
                ws.Cells[row, 4].Value = rnd.NextDouble() * 10000;
                ws.Cells[row, 5].FormulaR1C1 = "RC[-4]+RC[-1]";
                if (row % 10000==0)
                {
                    Console.WriteLine("{0:HH.mm.ss}\tWriting row {1}...", DateTime.Now, row);
                }
            }
            //Add a sum at the end
            ws.Cells[Rows + 1, 5].Formula = string.Format("Sum({0})", new ExcelAddress(1, 5, Rows, 5).Address);
            ws.Cells[Rows + 1, 5].Style.Font.Bold = true;
            ws.Cells[Rows + 1, 5].Style.Numberformat.Format = "#,##0.00";

            Console.WriteLine("{0:HH.mm.ss}\tWriting row {1}...", DateTime.Now, Rows);
            Console.WriteLine("{0:HH.mm.ss}\tFormating...", DateTime.Now);
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
            ws.Cells["E1"].Value = "For-\r\nmula";
            ws.View.FreezePanes(2, 1);

            using(var rng = ws.Cells["A1:E1"])
            {
                rng.Style.Font.Bold = true;
                rng.Style.Font.Color.SetColor(Color.White);
                rng.Style.WrapText = true;
                rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rng.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
            }

            ws.Column(1).Width = 10;
            ws.Column(2).Width = 15;
            ws.Column(3).Width = 12;
            ws.Column(4).Width = 10;
            ws.Column(5).Width = 20;

            //Now we set the sheetprotection and a password.
            ws.Cells[2, 3, Rows + 1, 4].Style.Locked = false;
            ws.Cells[2, 3, Rows + 1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[2, 3, Rows + 1, 4].Style.Fill.BackgroundColor.SetColor(Color.White);
            ws.Cells[1, 5, Rows + 2, 5].Style.Hidden = true;    //Hide the formula
            ws.Protection.SetPassword("EPPlus");

            ws.Select("C2");
            Console.WriteLine("{0:HH.mm.ss}\tSaving...", DateTime.Now);
            package.SaveAs(newFile);
            Console.WriteLine("{0:HH.mm.ss}\tDone!!", DateTime.Now);
        }
    }
}
