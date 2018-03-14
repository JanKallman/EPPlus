using System;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml.Utils;
using OfficeOpenXml;
using OfficeOpenXml.Utils.CompundDocument;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Dynamic;
using System.Linq;

namespace EPPlusTest
{
    /// <summary>
    /// Summary description for CompoundDoc
    /// </summary>
    [TestClass]
    public class CompoundDoc
    {
        [TestMethod, Ignore]
        public void Read()
        {
           //var doc = File.ReadAllBytes(@"c:\temp\vbaProject.bin");
           var doc = File.ReadAllBytes(@"c:\temp\vba.bin");
           var cd = new CompoundDocumentFile(doc);
           var ms = new MemoryStream();
           cd.Write(ms);
           printitems(cd.RootItem);
           File.WriteAllBytes(@"c:\temp\vba.bin", ms.ToArray());
        }

        private void printitems(CompoundDocumentItem item)
        {
            File.AppendAllText(@"c:\temp\items.txt", item.Name+ "\t");            
            foreach(var c in item.Children)
            {
                printitems(c);
            }
        }

        [TestMethod, Ignore ]
        public void ReadEncLong()
        {
            var doc=File.ReadAllBytes(@"c:\temp\EncrDocRead.xlsx");
            var cd = new CompoundDocumentFile(doc);
            var ms = new MemoryStream();
            cd.Write(ms);

            File.WriteAllBytes(@"c:\temp\vba.xlsx", ms.ToArray());
        }
        
        [TestMethod, Ignore]
        public void ReadVba()
        {
            var p = new ExcelPackage(new FileInfo(@"c:\temp\pricecheck.xlsm"));
            var vba = p.Workbook.VbaProject;
            p.SaveAs(new FileInfo(@"c:\temp\pricecheckSaved.xlsm"));
        }
        
        [TestMethod]
        public void Issue131()
        {
            var fi = TempFile.GenerateFile
                .WithFileExtention("xlsm");
               
            var package = new ExcelPackage(fi);
            var sheets = Enumerable.Range(1, 500)
                .Select(x => $"Sheet{x}");
            foreach (var sheet in sheets)
                package.Workbook.Worksheets.Add(sheet);

            package.Workbook.CreateVBAProject();
            package.Workbook.VbaProject.Modules.AddModule("Module1").Code
                = "\r\nPublic Sub SayHello()\r\nMsgBox(\"Hello\")\r\nEnd Sub\r\n";

            package.Save();
        }

        
        [TestMethod, Ignore]
        public void Sample7EncrLargeTest()
        {
            int Rows = 1000000;
            int colMult = 20;

            var newFile = TempFile.GenerateFile.WithFileExtention("xlsx");
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
                    for (int c = 0; c < colMult * 5; c += 5)
                    {
                        ws.SetValue(row, 1 + c, row);                               //The SetValue method is a little bit faster than using the Value property
                        ws.SetValue(row, 2 + c, string.Format("Row {0}", row));
                        ws.SetValue(row, 3 + c, DateTime.Today.AddDays(row));
                        ws.SetValue(row, 4 + c, rnd.NextDouble() * 10000); 
                     }
                }
                var endC = colMult * 5;
                ws.Cells[1, endC, Rows, endC].FormulaR1C1 = "RC[-4]+RC[-1]";

                //Add a sum at the end
                ws.Cells[Rows + 1, endC].Formula = string.Format("Sum({0})", new ExcelAddress(1, 5, Rows, 5).Address);
                ws.Cells[Rows + 1, endC].Style.Font.Bold = true;
                ws.Cells[Rows + 1, endC].Style.Numberformat.Format = "#,##0.00";

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

                //Calculate (Commented away thisk, it was a bit time consuming... /MA)
                // Console.WriteLine("{0:HH.mm.ss}\tCalculate formulas...", DateTime.Now);
                // ws.Calculate();

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
                package.Encryption.IsEncrypted = true;
                package.SaveAs(newFile);
            }
            Console.WriteLine("{0:HH.mm.ss}\tDone!!", DateTime.Now);
        }
        
        [TestMethod, Ignore]
        public void ReadPerfTest()
        {
            var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\sample7compdoctest.xlsx"), "");
            //var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\sample7compdoctest_4.5.xlsx"), "");
            //var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\sample7compdoctest.310k.xlsx"), "");
        }

        [TestMethod, Ignore]
        public void ReadVbaIssue107()
        {
            //var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\report.xlsm"));
            //var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\report411.xlsm"));
            var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\sample7.xlsx"),"");
            var vba = p.Workbook.VbaProject;
        }
        
    }
}
