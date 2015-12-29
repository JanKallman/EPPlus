using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.Style;
using System.Data;
using OfficeOpenXml.Table;
using Rhino.Mocks.Constraints;
using System.Collections.Generic;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlusTest
{
    [TestClass]
    public class Issues
    {
        [TestInitialize]
        public void Initialize()
        {
            if (!Directory.Exists(@"c:\Temp"))
            {
                Directory.CreateDirectory(@"c:\Temp");
            }
            if (!Directory.Exists(@"c:\Temp\bug"))
            {
                Directory.CreateDirectory(@"c:\Temp\bug");
            }
        }
        [TestMethod, Ignore]
        public void Issue15052()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("test");
            ws.Cells["A1:A4"].Value = 1;
            ws.Cells["B1:B4"].Value = 2;

            ws.Cells[1, 1, 4, 1]
                        .Style.Numberformat.Format = "#,##0.00;[Red]-#,##0.00";

            ws.Cells[1, 2, 5, 2]
                                    .Style.Numberformat.Format = "#,##0;[Red]-#,##0";

            p.SaveAs(new FileInfo(@"c:\temp\style.xlsx"));
        }
        [TestMethod]
        public void Issue15041()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells["A1"].Value = 202100083;
                ws.Cells["A1"].Style.Numberformat.Format = "00\\.00\\.00\\.000\\.0";
                Assert.AreEqual("02.02.10.008.3", ws.Cells["A1"].Text);
                ws.Dispose();
            }
        }
        [TestMethod]
        public void Issue15031()
        {
            var d = OfficeOpenXml.Utils.ConvertUtil.GetValueDouble(new TimeSpan(35, 59, 1));
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells["A1"].Value = d;
                ws.Cells["A1"].Style.Numberformat.Format = "[t]:mm:ss";
                ws.Dispose();
            }
        }
        [TestMethod]
        public void Issue15022()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells.AutoFitColumns();
                ws.Cells["A1"].Style.Numberformat.Format = "0";
                ws.Cells.AutoFitColumns();
            }
        }
        [TestMethod]
        public void Issue15056()
        {
            var path = @"C:\temp\output.xlsx";
            var file = new FileInfo(path);
            file.Delete();
            using (var ep = new ExcelPackage(file))
            {
                var s = ep.Workbook.Worksheets.Add("test");
                s.Cells["A1:A2"].Formula = ""; // or null, or non-empty whitespace, with same result
                ep.Save();
            }

        }
        [Ignore]
        [TestMethod]
        public void Issue15058()
        {
            System.IO.FileInfo newFile = new System.IO.FileInfo(@"C:\Temp\output.xlsx");
            ExcelPackage excelP = new ExcelPackage(newFile);
            ExcelWorksheet ws = excelP.Workbook.Worksheets[1];
        }
        [Ignore]
        [TestMethod]
        public void Issue15063()
        {
            System.IO.FileInfo newFile = new System.IO.FileInfo(@"C:\Temp\bug\TableFormula.xlsx");
            ExcelPackage excelP = new ExcelPackage(newFile);
            ExcelWorksheet ws = excelP.Workbook.Worksheets[1];
            ws.Calculate();
        }
        [Ignore]
        [TestMethod]
        public void Issue15112()
        {
            System.IO.FileInfo case1 = new System.IO.FileInfo(@"c:\temp\bug\src\src\DeleteRowIssue\Template.xlsx");
            var p = new ExcelPackage(case1);
            var first = p.Workbook.Worksheets[1];
            first.DeleteRow(5);
            p.SaveAs(new System.IO.FileInfo(@"c:\temp\bug\DeleteCol_case1.xlsx"));

            var case2 = new System.IO.FileInfo(@"c:\temp\bug\src2\DeleteRowIssue\Template.xlsx");
            p = new ExcelPackage(case2);
            first = p.Workbook.Worksheets[1];
            first.DeleteRow(5);
            p.SaveAs(new System.IO.FileInfo(@"c:\temp\bug\DeleteCol_case2.xlsx"));
        }

        [Ignore]
        [TestMethod]
        public void Issue15118()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bugOutput.xlsx"), new FileInfo(@"c:\temp\bug\DeleteRowIssue\Template.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];
                worksheet.Cells[9, 6, 9, 7].Merge = true;
                worksheet.Cells[9, 8].Merge = false;

                worksheet.DeleteRow(5);
                worksheet.DeleteColumn(5);
                worksheet.DeleteColumn(5);
                worksheet.DeleteColumn(5);
                worksheet.DeleteColumn(5);

                package.Save();
            }
        }
        [Ignore]
        [TestMethod]
        public void Issue15109()
        {
            System.IO.FileInfo newFile = new System.IO.FileInfo(@"C:\Temp\bug\test01.xlsx");
            ExcelPackage excelP = new ExcelPackage(newFile);
            ExcelWorksheet ws = excelP.Workbook.Worksheets[1];
            Assert.AreEqual("A1:Z75", ws.Dimension.Address);
            excelP.Dispose();

            newFile = new System.IO.FileInfo(@"C:\Temp\bug\test02.xlsx");
            excelP = new ExcelPackage(newFile);
            ws = excelP.Workbook.Worksheets[1];
            Assert.AreEqual("A1:AF501", ws.Dimension.Address);
            excelP.Dispose();

            newFile = new System.IO.FileInfo(@"C:\Temp\bug\test03.xlsx");
            excelP = new ExcelPackage(newFile);
            ws = excelP.Workbook.Worksheets[1];
            Assert.AreEqual("A1:AD406", ws.Dimension.Address);
            excelP.Dispose();
        }
        [Ignore]
        [TestMethod]
        public void Issue15120()
        {
            var p = new ExcelPackage(new System.IO.FileInfo(@"C:\Temp\bug\pp.xlsx"));
            ExcelWorksheet ws = p.Workbook.Worksheets["tum_liste"];
            ExcelWorksheet wPvt = p.Workbook.Worksheets.Add("pvtSheet");
            var pvSh = wPvt.PivotTables.Add(wPvt.Cells["B5"], ws.Cells[ws.Dimension.Address.ToString()], "pvtS");

            //p.Save();
        }
        [TestMethod]
        public void Issue15113()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Cells["A1"].Value = " Performance Update";
            ws.Cells["A1:H1"].Merge = true;
            ws.Cells["A1:H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
            ws.Cells["A1:H1"].Style.Font.Size = 14;
            ws.Cells["A1:H1"].Style.Font.Color.SetColor(Color.Red);
            ws.Cells["A1:H1"].Style.Font.Bold = true;
            p.SaveAs(new FileInfo(@"c:\temp\merge.xlsx"));
        }
        [TestMethod]
        public void Issue15141()
        {
            using (ExcelPackage package = new ExcelPackage())
            using (ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Test"))
            {
                sheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                sheet.Cells[1, 1, 1, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sheet.Cells[1, 5, 2, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                ExcelColumn column = sheet.Column(3); // fails with exception
            }
        }
        [TestMethod, Ignore]
        public void Issue15145()
        {
            using (ExcelPackage p = new ExcelPackage(new System.IO.FileInfo(@"C:\Temp\bug\ColumnInsert.xlsx")))
            {
                ExcelWorksheet ws = p.Workbook.Worksheets[1];
                ws.InsertColumn(12, 3);
                ws.InsertRow(30, 3);
                ws.DeleteRow(31, 1);
                ws.DeleteColumn(7, 1);
                p.SaveAs(new System.IO.FileInfo(@"C:\Temp\bug\InsertCopyFail.xlsx"));
            }
        }
        [TestMethod, Ignore]
        public void Issue15150()
        {
            var template = new FileInfo(@"c:\temp\bug\ClearIssue.xlsx");
            const string output = @"c:\temp\bug\ClearIssueSave.xlsx";

            using (var pck = new ExcelPackage(template, false))
            {
                var ws = pck.Workbook.Worksheets[1];
                ws.Cells["A2:C3"].Value = "Test";
                var c = ws.Cells["B2:B3"];
                c.Clear();

                pck.SaveAs(new FileInfo(output));
            }
        }

        [TestMethod, Ignore]
        public void Issue15146()
        {
            var template = new FileInfo(@"c:\temp\bug\CopyFail.xlsx");
            const string output = @"c:\temp\bug\CopyFail-Save.xlsx";

            using (var pck = new ExcelPackage(template, false))
            {
                var ws = pck.Workbook.Worksheets[3];

                //ws.InsertColumn(3, 1);
                CustomColumnInsert(ws, 3, 1);

                pck.SaveAs(new FileInfo(output));
            }
        }

        private static void CustomColumnInsert(ExcelWorksheet ws, int column, int columns)
        {
            var source = ws.Cells[1, column, ws.Dimension.End.Row, ws.Dimension.End.Column];
            var dest = ws.Cells[1, column + columns, ws.Dimension.End.Row, ws.Dimension.End.Column + columns];
            source.Copy(dest);
        }
        [TestMethod]
        public void Issue15123()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            using (var dt = new DataTable())
            {
                dt.Columns.Add("String", typeof(string));
                dt.Columns.Add("Int", typeof(int));
                dt.Columns.Add("Bool", typeof(bool));
                dt.Columns.Add("Double", typeof(double));
                dt.Columns.Add("Date", typeof(DateTime));

                var dr = dt.NewRow();
                dr[0] = "Row1";
                dr[1] = 1;
                dr[2] = true;
                dr[3] = 1.5;
                dr[4] = new DateTime(2014, 12, 30);
                dt.Rows.Add(dr);

                dr = dt.NewRow();
                dr[0] = "Row2";
                dr[1] = 2;
                dr[2] = false;
                dr[3] = 2.25;
                dr[4] = new DateTime(2014, 12, 31);
                dt.Rows.Add(dr);

                ws.Cells["A1"].LoadFromDataTable(dt, true);
                ws.Cells["D2:D3"].Style.Numberformat.Format = "(* #,##0.00);_(* (#,##0.00);_(* \"-\"??_);(@)";

                ws.Cells["E2:E3"].Style.Numberformat.Format = "mm/dd/yyyy";
                ws.Cells.AutoFitColumns();
                Assert.AreNotEqual(ws.Cells[2, 5].Text, "");
            }
        }

        [TestMethod]
        public void Issue15128()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Cells["A1"].Value = 1;
            ws.Cells["B1"].Value = 2;
            ws.Cells["B2"].Formula = "A1+$B$1";
            ws.Cells["C1"].Value = "Test";
            ws.Cells["A1:B2"].Copy(ws.Cells["C1"]);
            ws.Cells["B2"].Copy(ws.Cells["D1"]);
            p.SaveAs(new FileInfo(@"c:\temp\bug\copy.xlsx"));
        }

        [TestMethod]
        public void IssueMergedCells()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Cells["A1:A5,C1:C8"].Merge = true;
            ws.Cells["C1:C8"].Merge = false;
            ws.Cells["A1:A8"].Merge = false;
            p.Dispose();
        }
        [Ignore]
        [TestMethod]
        public void Issue15158()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\Output.xlsx"), new FileInfo(@"C:\temp\bug\DeleteColFormula\FormulasIssue\demo.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                ExcelWorksheet worksheet = workBook.Worksheets[1];

                //string column = ColumnIndexToColumnLetter(28);
                worksheet.DeleteColumn(28);

                if (worksheet.Cells["AA19"].Formula != "")
                {
                    throw new Exception("this cell should not have formula");
                }

                package.Save();
            }
        }

        public class cls1
        {
            public int prop1 { get; set; }
        }

        public class cls2 : cls1
        {
            public string prop2 { get; set; }
        }
        [TestMethod]
        public void LoadFromColIssue()
        {
            var l = new List<cls1>();

            var c2 = new cls2() { prop1 = 1, prop2 = "test1" };
            l.Add(c2);

            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Test");

            ws.Cells["A1"].LoadFromCollection(l, true, TableStyles.Light16, BindingFlags.Instance | BindingFlags.Public,
                new MemberInfo[] { typeof(cls2).GetProperty("prop2") });
        }

        [TestMethod]
        public void Issue15168()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Test");
                ws.Cells[1, 1].Value = "A1";
                ws.Cells[2, 1].Value = "A2";

                ws.Cells[2, 1].Value = ws.Cells[1, 1].Value;
                Assert.AreEqual("A1", ws.Cells[1, 1].Value);
            }
        }
        [Ignore]
        [TestMethod]
        public void Issue15159()
        {
            var fs = new FileStream(@"C:\temp\bug\DeleteColFormula\FormulasIssue\demo.xlsx", FileMode.OpenOrCreate);
            using (var package = new OfficeOpenXml.ExcelPackage(fs))
            {
                package.Save();
            }
            fs.Seek(0, SeekOrigin.Begin);
            var fs2 = fs;
        }
        [TestMethod]
        public void Issue15179()
        {
            using (var package = new OfficeOpenXml.ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("MergeDeleteBug");
                ws.Cells["E3:F3"].Merge = true;
                ws.Cells["E3:F3"].Merge = false;
                ws.DeleteRow(2, 6);
                ws.Cells["A1"].Value = 0;
                var s = ws.Cells["A1"].Value.ToString();

            }
        }
        [Ignore]
        [TestMethod]
        public void Issue15169()
        {
            FileInfo fileInfo = new FileInfo(@"C:\temp\bug\issue\input.xlsx");

            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            {
                string sheetName = "Labour Costs";

                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[sheetName];
                excelPackage.Workbook.Worksheets.Delete(ws);

                ws = excelPackage.Workbook.Worksheets.Add(sheetName);

                excelPackage.SaveAs(new FileInfo(@"C:\temp\bug\issue\output2.xlsx"));
            }
        }
        [Ignore]
        [TestMethod]
        public void Issue15172()
        {
            FileInfo fileInfo = new FileInfo(@"C:\temp\bug\book2.xlsx");

            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            {
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];

                Assert.AreEqual("IF($R10>=X$2,1,0)", ws.Cells["X10"].Formula);
                ws.Calculate();
                Assert.AreEqual(0D, ws.Cells["X10"].Value);
            }
        }
        [Ignore]
        [TestMethod]
        public void Issue15174()
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(@"C:\temp\bug\MyTemplate.xlsx")))
            {
                package.Workbook.Worksheets[1].Column(2).Style.Numberformat.Format = "dd/mm/yyyy";

                package.SaveAs(new FileInfo(@"C:\temp\bug\MyTemplate2.xlsx"));
            }
        }
        [Ignore]
        [TestMethod]
        public void PictureIssue()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Drawings.AddPicture("Test", new FileInfo(@"c:\temp\bug\2152228.jpg"));
            p.SaveAs(new FileInfo(@"c:\temp\bug\pic.xlsx"));
        }

        [Ignore]
        [TestMethod]
        public void Issue14988()
        {
            var guid = Guid.NewGuid().ToString("N");
            using (var outputStream = new FileStream(@"C:\temp\" + guid + ".xlsx", FileMode.Create))
            {
                using (var inputStream = new FileStream(@"C:\temp\bug2.xlsx", FileMode.Open))
                {
                    using (var package = new ExcelPackage(outputStream, inputStream, "Test"))
                    {
                        var ws = package.Workbook.Worksheets.Add("Test empty");
                        ws.Cells["A1"].Value = "Test";
                        package.Encryption.Password = "Test2";
                        package.Save();
                        //package.SaveAs(new FileInfo(@"c:\temp\test2.xlsx"));
                    }
                }
            }
        }

        [TestMethod, Ignore]
        public void Issue15173_1()
        {
            using (var pck = new ExcelPackage(new FileInfo(@"c:\temp\EPPlusIssues\Excel01.xlsx")))
            {
                var sw = new Stopwatch();
                //pck.Workbook.FormulaParser.Configure(x => x.AttachLogger(LoggerFactory.CreateTextFileLogger(new FileInfo(@"c:\Temp\log1.txt"))));
                sw.Start();
                var ws = pck.Workbook.Worksheets.First();
                pck.Workbook.Calculate();
                Assert.AreEqual("20L2300", ws.Cells["F4"].Value);
                Assert.AreEqual("20K2E01", ws.Cells["F5"].Value);
                var f7Val = pck.Workbook.Worksheets["MODELLO-TIPO PANNELLO"].Cells["F7"].Value;
                Assert.AreEqual(13.445419, Math.Round((double)f7Val, 6));
                sw.Stop();
                Console.WriteLine(sw.Elapsed.TotalSeconds); // approx. 10 seconds

            }
        }

        [TestMethod, Ignore]
        public void Issue15173_2()
        {
            using (var pck = new ExcelPackage(new FileInfo(@"c:\temp\EPPlusIssues\Excel02.xlsx")))
            {
                var sw = new Stopwatch();
                pck.Workbook.FormulaParser.Configure(x => x.AttachLogger(LoggerFactory.CreateTextFileLogger(new FileInfo(@"c:\Temp\log1.txt"))));
                sw.Start();
                var ws = pck.Workbook.Worksheets.First();
                //ws.Calculate();
                pck.Workbook.Calculate();
                Assert.AreEqual("20L2300", ws.Cells["F4"].Value);
                Assert.AreEqual("20K2E01", ws.Cells["F5"].Value);
                sw.Stop();
                Console.WriteLine(sw.Elapsed.TotalSeconds); // approx. 10 seconds

            }
        }
        [Ignore]
        [TestMethod]
        public void Issue15154()
        {
            Directory.EnumerateFiles(@"c:\temp\bug\ConstructorInvokationNotThreadSafe\").AsParallel().ForAll(file =>
            {
                //lock (_lock)
                //{
                using (var package = new ExcelPackage(new FileStream(file, FileMode.Open)))
                {
                    package.Workbook.Worksheets[1].Cells[1, 1].Value = file;
                    package.SaveAs(new FileInfo(@"c:\temp\bug\ConstructorInvokationNotThreadSafe\new\" + new FileInfo(file).Name));
                }
                //}
            });

        }
        [Ignore]
        [TestMethod]
        public void Issue15188()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("test");
                worksheet.Column(6).Style.Numberformat.Format = "mm/dd/yyyy";
                worksheet.Column(7).Style.Numberformat.Format = "mm/dd/yyyy";
                worksheet.Column(8).Style.Numberformat.Format = "mm/dd/yyyy";
                worksheet.Column(10).Style.Numberformat.Format = "mm/dd/yyyy";

                worksheet.Cells[2, 6].Value = DateTime.Today;
                string a = worksheet.Cells[2, 6].Text;
                Assert.AreEqual(DateTime.Today.ToString("MM/dd/yyyy"), a);
            }
        }
        [TestMethod, Ignore]
        public void Issue15194()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\i15194-Save.xlsx"), new FileInfo(@"c:\temp\bug\I15194.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                worksheet.Cells["E3:F3"].Merge = false;

                worksheet.DeleteRow(2, 6);

                package.Save();
            }
        }
        [TestMethod, Ignore]
        public void Issue15195()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\i15195_Save.xlsx"), new FileInfo(@"c:\temp\bug\i15195.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                worksheet.InsertColumn(8, 2);

                package.Save();
            }
        }
        [TestMethod, Ignore]
        public void Issue14788()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\i15195_Save.xlsx"), new FileInfo(@"c:\temp\bug\GetWorkSheetXmlBad.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                worksheet.InsertColumn(8, 2);

                package.Save();
            }
        }
        [TestMethod, Ignore]
        public void Issue15167()
        {
            FileInfo fileInfo = new FileInfo(@"c:\temp\bug\Draw\input.xlsx");

            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            {
                string sheetName = "Board pack";

                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[sheetName];
                excelPackage.Workbook.Worksheets.Delete(ws);

                ws = excelPackage.Workbook.Worksheets.Add(sheetName);

                excelPackage.SaveAs(new FileInfo(@"c:\temp\bug\output.xlsx"));
            }
        }
        [TestMethod, Ignore]
        public void Issue15198()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\Output.xlsx"), new FileInfo(@"c:\temp\bug\demo.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                worksheet.DeleteRow(12);

                package.Save();
            }
        }
        [TestMethod, Ignore]
        public void Issue13492()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\Bug13492.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                var rt = worksheet.Cells["K31"].RichText.Text;

                package.Save();
            }
        }
        [TestMethod, Ignore]
        public void Issue14966()
        {
            using (var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\ssis\FileFromReportingServer2012.xlsx")))
                package.SaveAs(new FileInfo(@"c:\temp\bug\ssis\Corrupted.xlsx"));
        }
        [TestMethod, Ignore]
        public void Issue15200()
        {
            File.Copy(@"C:\temp\bug\EPPlusRangeCopyTest\EPPlusRangeCopyTest\input.xlsx", @"C:\temp\bug\EPPlusRangeCopyTest\EPPlusRangeCopyTest\output.xlsx", true);

            using (var p = new ExcelPackage(new FileInfo(@"C:\temp\bug\EPPlusRangeCopyTest\EPPlusRangeCopyTest\output.xlsx")))
            {
                var sheet = p.Workbook.Worksheets.First();

                var sourceRange = sheet.Cells[1, 1, 1, 2];
                var resultRange = sheet.Cells[3, 1, 3, 2];
                sourceRange.Copy(resultRange);

                sourceRange = sheet.Cells[1, 1, 1, 7];
                resultRange = sheet.Cells[5, 1, 5, 7];
                sourceRange.Copy(resultRange);  // This throws System.ArgumentException: Can't merge and already merged range

                sourceRange = sheet.Cells[1, 1, 1, 7];
                resultRange = sheet.Cells[7, 3, 7, 7];
                sourceRange.Copy(resultRange);  // This throws System.ArgumentException: Can't merge and already merged range

                p.Save();
            }
        }
        [TestMethod]
        public void Issue15212()
        {
            var s = "_(\"R$ \"* #,##0.00_);_(\"R$ \"* (#,##0.00);_(\"R$ \"* \"-\"??_);_(@_) )";
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("StyleBug");
                ws.Cells["A1"].Value = 5698633.64;
                ws.Cells["A1"].Style.Numberformat.Format = s;
                var t = ws.Cells["A1"].Text;
            }
        }
        [TestMethod, Ignore]
        public void Issue15213()
        {
            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\ExcelClearDemo\exceltestfile.xlsx")))
            {
                foreach (var ws in p.Workbook.Worksheets)
                {
                    ws.Cells[1023, 1, ws.Dimension.End.Row - 2, ws.Dimension.End.Column].Clear();
                    Assert.AreNotEqual(ws.Dimension, null);
                }
                foreach (var cell in p.Workbook.Worksheets[2].Cells)
                {
                    Console.WriteLine(cell);
                }
                p.SaveAs(new FileInfo(@"c:\temp\bug\ExcelClearDemo\exceltestfile-save.xlsx"));
            }
        }

        [TestMethod, Ignore]
        public void Issuer15217()
        {

            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\FormatRowCol.xlsx")))
            {
                var ws = p.Workbook.Worksheets.Add("fmt");
                ws.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Row(1).Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                ws.Cells["A1:B2"].Value = 1;
                ws.Column(1).Style.Numberformat.Format = "yyyy-mm-dd hh:mm";
                ws.Column(2).Style.Numberformat.Format = "#,##0";
                p.Save();
            }
        }
        [TestMethod, Ignore]
        public void Issuer15228()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("colBug");

                var col = ws.Column(7);
                col.ColumnMax = 8;
                col.Hidden = true;

                var col8 = ws.Column(8);
                Assert.AreEqual(true, col8.Hidden);
            }
        }

        [TestMethod, Ignore]
        public void Issue15234()
        {
            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\merge2\input.xlsx")))
            {
                var sheet = p.Workbook.Worksheets.First();

                var sourceRange = sheet.Cells["1:4"];

                sheet.InsertRow(5, 4);

                var resultRange = sheet.Cells["5:8"];
                sourceRange.Copy(resultRange);

                p.Save();
            }
        }
        [TestMethod]
        /**** Pivottable issue ****/
        public void Issue()
        {
            DirectoryInfo outputDir = new DirectoryInfo(@"c:\ExcelPivotTest");
            FileInfo MyFile = new FileInfo(@"c:\temp\bug\pivottable.xlsx");
            LoadData(MyFile);
            BuildPivotTable1(MyFile);
            BuildPivotTable2(MyFile);
        }

        private void LoadData(FileInfo MyFile)
        {
            if (MyFile.Exists)
            {
                MyFile.Delete();  // ensures we create a new workbook
            }

            using (ExcelPackage EP = new ExcelPackage(MyFile))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet wsData = EP.Workbook.Worksheets.Add("Data");
                //Add the headers
                wsData.Cells[1, 1].Value = "INVOICE_DATE";
                wsData.Cells[1, 2].Value = "TOTAL_INVOICE_PRICE";
                wsData.Cells[1, 3].Value = "EXTENDED_PRICE_VARIANCE";
                wsData.Cells[1, 4].Value = "AUDIT_LINE_STATUS";
                wsData.Cells[1, 5].Value = "RESOLUTION_STATUS";
                wsData.Cells[1, 6].Value = "COUNT";

                //Add some items...
                wsData.Cells["A2"].Value = Convert.ToDateTime("04/2/2012");
                wsData.Cells["B2"].Value = 33.63;
                wsData.Cells["C2"].Value = (-.87);
                wsData.Cells["D2"].Value = "Unfavorable Price Variance";
                wsData.Cells["E2"].Value = "Pending";
                wsData.Cells["F2"].Value = 1;

                wsData.Cells["A3"].Value = Convert.ToDateTime("04/2/2012");
                wsData.Cells["B3"].Value = 43.14;
                wsData.Cells["C3"].Value = (-1.29);
                wsData.Cells["D3"].Value = "Unfavorable Price Variance";
                wsData.Cells["E3"].Value = "Pending";
                wsData.Cells["F3"].Value = 1;

                wsData.Cells["A4"].Value = Convert.ToDateTime("11/8/2011");
                wsData.Cells["B4"].Value = 55;
                wsData.Cells["C4"].Value = (-2.87);
                wsData.Cells["D4"].Value = "Unfavorable Price Variance";
                wsData.Cells["E4"].Value = "Pending";
                wsData.Cells["F4"].Value = 1;

                wsData.Cells["A5"].Value = Convert.ToDateTime("11/8/2011");
                wsData.Cells["B5"].Value = 38.72;
                wsData.Cells["C5"].Value = (-5.00);
                wsData.Cells["D5"].Value = "Unfavorable Price Variance";
                wsData.Cells["E5"].Value = "Pending";
                wsData.Cells["F5"].Value = 1;

                wsData.Cells["A6"].Value = Convert.ToDateTime("3/4/2011");
                wsData.Cells["B6"].Value = 77.44;
                wsData.Cells["C6"].Value = (-1.55);
                wsData.Cells["D6"].Value = "Unfavorable Price Variance";
                wsData.Cells["E6"].Value = "Pending";
                wsData.Cells["F6"].Value = 1;

                wsData.Cells["A7"].Value = Convert.ToDateTime("3/4/2011");
                wsData.Cells["B7"].Value = 127.55;
                wsData.Cells["C7"].Value = (-10.50);
                wsData.Cells["D7"].Value = "Unfavorable Price Variance";
                wsData.Cells["E7"].Value = "Pending";
                wsData.Cells["F7"].Value = 1;

                using (var range = wsData.Cells[2, 1, 7, 1])
                {
                    range.Style.Numberformat.Format = "mm-dd-yy";
                }

                wsData.Cells.AutoFitColumns(0);
                EP.Save();
            }
        }
        private void BuildPivotTable1(FileInfo MyFile)
        {
            using (ExcelPackage ep = new ExcelPackage(MyFile))
            {

                var wsData = ep.Workbook.Worksheets["Data"];
                var totalRows = wsData.Dimension.Address;
                ExcelRange data = wsData.Cells[totalRows];

                var wsAuditPivot = ep.Workbook.Worksheets.Add("Pivot1");

                var pivotTable1 = wsAuditPivot.PivotTables.Add(wsAuditPivot.Cells["A7:C30"], data, "PivotAudit1");
                pivotTable1.ColumGrandTotals = true;
                var rowField = pivotTable1.RowFields.Add(pivotTable1.Fields["INVOICE_DATE"]);


                rowField.AddDateGrouping(eDateGroupBy.Years);
                var yearField = pivotTable1.Fields.GetDateGroupField(eDateGroupBy.Years);
                yearField.Name = "Year";

                var rowField2 = pivotTable1.RowFields.Add(pivotTable1.Fields["AUDIT_LINE_STATUS"]);

                var TotalSpend = pivotTable1.DataFields.Add(pivotTable1.Fields["TOTAL_INVOICE_PRICE"]);
                TotalSpend.Name = "Total Spend";
                TotalSpend.Format = "$##,##0";


                var CountInvoicePrice = pivotTable1.DataFields.Add(pivotTable1.Fields["COUNT"]);
                CountInvoicePrice.Name = "Total Lines";
                CountInvoicePrice.Format = "##,##0";

                pivotTable1.DataOnRows = false;
                ep.Save();
                ep.Dispose();

            }

        }

        private void BuildPivotTable2(FileInfo MyFile)
        {
            using (ExcelPackage ep = new ExcelPackage(MyFile))
            {

                var wsData = ep.Workbook.Worksheets["Data"];
                var totalRows = wsData.Dimension.Address;
                ExcelRange data = wsData.Cells[totalRows];

                var wsAuditPivot = ep.Workbook.Worksheets.Add("Pivot2");

                var pivotTable1 = wsAuditPivot.PivotTables.Add(wsAuditPivot.Cells["A7:C30"], data, "PivotAudit2");
                pivotTable1.ColumGrandTotals = true;
                var rowField = pivotTable1.RowFields.Add(pivotTable1.Fields["INVOICE_DATE"]);


                rowField.AddDateGrouping(eDateGroupBy.Years);
                var yearField = pivotTable1.Fields.GetDateGroupField(eDateGroupBy.Years);
                yearField.Name = "Year";

                var rowField2 = pivotTable1.RowFields.Add(pivotTable1.Fields["AUDIT_LINE_STATUS"]);

                var TotalSpend = pivotTable1.DataFields.Add(pivotTable1.Fields["TOTAL_INVOICE_PRICE"]);
                TotalSpend.Name = "Total Spend";
                TotalSpend.Format = "$##,##0";


                var CountInvoicePrice = pivotTable1.DataFields.Add(pivotTable1.Fields["COUNT"]);
                CountInvoicePrice.Name = "Total Lines";
                CountInvoicePrice.Format = "##,##0";

                pivotTable1.DataOnRows = false;
                ep.Save();
                ep.Dispose();

            }

        }
        [TestMethod, Ignore]
        public string Issue15247(DirectoryInfo outputDir)
        {
            FileInfo templateFile = new FileInfo(outputDir.FullName + @"\diagonal.xlsx");
            FileInfo newFile = new FileInfo(outputDir.FullName + @"\sampleDiagonal.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(outputDir.FullName + @"\sampleDiagonal.xlsx");
            }
            using (ExcelPackage package = new ExcelPackage(newFile, templateFile))
            {
                //Open worksheet 1
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                Debug.WriteLine(worksheet.Cells["A1"].Style.Border.DiagonalUp);
                Debug.WriteLine(worksheet.Cells["A1"].Style.Border.DiagonalDown);


                //worksheet.Cells["A1"].Style.Border.DiagonalUp = true;
                //worksheet.Cells["A1"].Style.Border.DiagonalDown = true;

                // save our new workbook and we are done!
                package.Save();
            }

            return newFile.FullName;
        }
        [TestMethod, Ignore]
        public void issue15249()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\Boldtextcopy.xlsx")))
            {
                exfile.Workbook.Worksheets.Copy("sheet1", "copiedSheet");
                exfile.SaveAs(new FileInfo(@"c:\temp\bug\Boldtextcopy2.xlsx"));
            }
        }
        [TestMethod, Ignore]
        public void issue15300()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\headfootpic.xlsx")))
            {
                exfile.Workbook.Worksheets.Copy("sheet1", "copiedSheet");
                exfile.SaveAs(new FileInfo(@"c:\temp\bug\headfootpic_save.xlsx"));
            }

        }
        [TestMethod, Ignore]
        public void issue15295()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"C:\temp\bug\pivot issue\input.xlsx")))
            {
                exfile.SaveAs(new FileInfo(@"C:\temp\bug\pivot issue\pivotcoldup.xlsx"));
            }

        }
        [TestMethod, Ignore]
        public void issue15282()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"C:\temp\bug\pivottable-table.xlsx")))
            {
                exfile.SaveAs(new FileInfo(@"C:\temp\bug\pivot issue\pivottab-tab-save.xlsx"));
            }

        }

        [TestMethod, Ignore]
        public void Issues14699()
        {
            FileInfo newFile = new FileInfo(string.Format("c:\\temp\\bug\\EPPlus_Issue14699.xlsx", System.IO.Directory.GetCurrentDirectory()));
            OfficeOpenXml.ExcelPackage pkg = new ExcelPackage(newFile);
            ExcelWorksheet wksheet = pkg.Workbook.Worksheets.Add("Issue14699");
            // Initialize a small range
            for (int row = 1; row < 11; row++)
            {
                for (int col = 1; col < 11; col++)
                {
                    wksheet.Cells[row, col].Value = string.Format("{0}{1}", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[col - 1], row);
                }
            }
            wksheet.View.FreezePanes(3, 3);
            pkg.Save();

        }
        [TestMethod, Ignore]
        public void Issue15382()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\Text Run Issue.xlsx")))
            {
                exfile.SaveAs(new FileInfo(@"C:\temp\bug\inlinText.xlsx"));
            }
        }
        [TestMethod, Ignore]
        public void Issue15380()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\dotinname.xlsx")))
            {
                var v = exfile.Workbook.Worksheets["sheet1.3"].Names["Test.Name"].Value;
                Assert.AreEqual(v, 1);
            }
        }
        [TestMethod, Ignore]
        public void Issue15378()
        {
            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\bubble.xlsx")))
            {
                var c = p.Workbook.Worksheets[1].Drawings[0] as ExcelBubbleChart;
                var cs = c.Series[0] as ExcelBubbleChartSerie;
            }
        }
        [TestMethod]
        public void Issue15377()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("ws1");
                ws.Cells["A1"].Value = (double?)1;
                var v = ws.GetValue<double?>(1,1);
            }
        }
        [TestMethod]
        public void Issue15374()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RT");
                var r = ws.Cells["A1"];
                r.RichText.Text = "Cell 1";
                r["A2"].RichText.Add("Cell 2");
                p.SaveAs(new FileInfo(@"c:\temp\rt.xlsx"));
            }
        }
        [TestMethod]
        public void IssueTranslate()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Trans");
                ws.Cells["A1:A2"].Formula = "IF(1=1, \"A's B C\",\"D\") ";
                var fr = ws.Cells["A1:A2"].FormulaR1C1;
                ws.Cells["A1:A2"].FormulaR1C1=fr;
                Assert.AreEqual("IF(1=1, \"A's B C\",\"D\")", ws.Cells["A2"].Formula);
            }
        }
        [TestMethod]
        public void Issue   ()
        {
            using (var p = new ExcelPackage())
            {
                var workSheet = p.Workbook.Worksheets.Add("styleerror");
                workSheet.Cells["F:G"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells["F:G"].Style.Fill.BackgroundColor.SetColor(Color.Red);

                workSheet.Cells["A:A,C:C"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells["A:A,C:C"].Style.Fill.BackgroundColor.SetColor(Color.Red);

                //And then: 

                workSheet.Cells["A:H"].Style.Font.Color.SetColor(Color.Blue);

                workSheet.Cells["I:I"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells["I:I"].Style.Fill.BackgroundColor.SetColor(Color.Red);
                workSheet.Cells["I2"].Style.Fill.BackgroundColor.SetColor(Color.Green);
                workSheet.Cells["I4"].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                workSheet.Cells["I9"].Style.Fill.BackgroundColor.SetColor(Color.Pink);

                workSheet.InsertColumn(2, 2, 9);

                p.SaveAs(new FileInfo(@"c:\temp\styleerror.xlsx"));
            }
        }
    }
}