using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.ConditionalFormatting;
using System.Threading;
using System.Drawing;
namespace EPPlusTest
{
    [TestClass]
    public class ReadTemplate //: TestBase
    {
        //[ClassInitialize()]
        //public static void ClassInit(TestContext testContext)
        //{
        //    //InitBase();
        //}
        //[ClassCleanup()]
        //public static void ClassCleanup()
        //{
        //    //SaveWorksheet("Worksheet.xlsx");
        //}
        [TestMethod]
        public void ReadBlankStream()
        {
            MemoryStream stream = new MemoryStream();
            using (ExcelPackage pck = new ExcelPackage(stream))
            {
                var ws = pck.Workbook.Worksheets.Add("Perf");
                pck.SaveAs(stream);
            }
            stream.Close();
        }
        [Ignore]
        [TestMethod]
        public void ReadBug()
        {
            var file = new FileInfo(@"c:\temp\Adenoviridae Protocol.xlsx");
            using (ExcelPackage pck = new ExcelPackage(file))
            {
                pck.Workbook.Worksheets[1].Cells["G4"].Value=12;
                pck.SaveAs(new FileInfo(@"c:\temp\Adenoviridae Protocol2.xlsx"));
            }
        }
        [Ignore]
        [TestMethod]
        public void ReadBug3()
        {
            ExcelPackage xlsPack = new ExcelPackage(new FileInfo(@"c:\temp\billing_template.xlsx"));
            ExcelWorkbook xlsWb = xlsPack.Workbook;
            ExcelWorksheet xlsSheet = xlsWb.Worksheets["Billing"];
        }
        [Ignore]
        [TestMethod]
        public void ReadBug2()
        {
            var file = new FileInfo(@"c:\temp\book2.xlsx");
            using (ExcelPackage pck = new ExcelPackage(file))
            {
                Assert.AreEqual("Good", pck.Workbook.Worksheets[1].Cells["A1"].StyleName);
                Assert.AreEqual("Good 2", pck.Workbook.Worksheets[1].Cells["C1"].StyleName);
                Assert.AreEqual("Note", pck.Workbook.Worksheets[1].Cells["G11"].StyleName);
                pck.SaveAs(new FileInfo(@"c:\temp\Adenoviridae Protocol2.xlsx"));
            }
        }
        [Ignore]
        [TestMethod]
        public void CondFormatDataValBug()
        {            
            var file = new FileInfo(@"c:\temp\condi.xlsx");
            using (ExcelPackage pck = new ExcelPackage(file))
            {
                var dv = pck.Workbook.Worksheets[1].Cells["A1"].DataValidation.AddIntegerDataValidation();
                dv.Formula.Value = 1;
                dv.Formula2.Value = 4;
                dv.Operator = OfficeOpenXml.DataValidation.ExcelDataValidationOperator.equal;
                pck.SaveAs(new FileInfo(@"c:\temp\condi2.xlsx"));
            }
        }
        [Ignore]
        [TestMethod]
        public void InternalZip()
        {
            //var file = @"c:\temp\condi.xlsx";
            //using (ExcelPackage pck = new ExcelPackage(file))
            //{
            //}
        }
        [Ignore]
        [TestMethod]
        public void ReadBug4()
        {
            var lines = new List<string>();
            var package = new ExcelPackage(new FileInfo(@"c:\temp\test.xlsx"));

            ExcelWorkbook workBook = package.Workbook;
            if (workBook != null)
            {
            if (workBook.Worksheets.Count > 0) //fails on this line
            {
            // Get the first worksheet
            ExcelWorksheet currentWorksheet = workBook.Worksheets.First();

            var rowCount = 1;
            var lastRow = currentWorksheet.Dimension.End.Row;
            var lastColumn = currentWorksheet.Dimension.End.Column;
            while (rowCount <= lastRow)
            {
            var columnCount = 1;
            var line = "";
            while (columnCount <= lastColumn)
            {
            line += currentWorksheet.Cells[rowCount, columnCount].Value + "|";
            columnCount++;
            }
            lines.Add(line);
            rowCount++;
            }
            }
            }
        }
        [Ignore]
        [TestMethod]
        public void ReadBug5()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\2.9 bugs\protect.xlsx"));

            package.Workbook.Worksheets[1].Protection.AllowInsertColumns = true;
            package.Workbook.Worksheets[1].Protection.SetPassword("test");
            package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\protectnew.xlsx"));
        }
        [Ignore]
        [TestMethod]
        public void ReadBug6()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\2.9 bugs\outofrange\error.xlsx"));

            package.Workbook.Worksheets[1].Protection.AllowInsertColumns = true;
            package.Workbook.Worksheets[1].Protection.SetPassword("test");
            package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\error.xlsx"));
        }
        [Ignore]
        [TestMethod]
        public void ReadBug7()
        {
            var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("test");
            using (var rng = ws.Cells["A1"])
            {
                var rt1 = rng.RichText.Add("TEXT1\r\n");
                rt1.Bold = true;
                rng.Style.WrapText = true;
                var rt2=rng.RichText.Add("TEXT2");
                rt2.Bold = false;
            }
            
            package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\error.xlsx"));
        }
        [Ignore]
        [TestMethod]
        public void ReadBug8()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\2.9 bugs\bug\Genband SO CrossRef Phoenix.xlsx"));
            var ws = package.Workbook.Worksheets[1];
            using (var rng = ws.Cells["A1"])
            {
                var rt1 = rng.RichText.Add("TEXT1\r\n");
                rt1.Bold = true;
                rng.Style.WrapText = true;
                var rt2 = rng.RichText.Add("TEXT2");
                rt2.Bold = false;
            }

            package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\billing_template.xlsx.error"));
        }
        [Ignore]
        [TestMethod]
        public void ReadBug9()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\CovenantsCheckReportTemplate.xlsx"));
            var ws = package.Workbook.Worksheets[1];
            package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\new_t.xlsx"));
        }
        [Ignore]
        [TestMethod]
        public void ReadBug10()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\Model_graphes_MBW.xlsm"));

            var ws = package.Workbook.Worksheets["HTTP_data"];
            Assert.IsNotNull(ws.Cells["B4"].Style.Fill.BackgroundColor.Indexed);
            Assert.IsNotNull(ws.Cells["B5"].Style.Fill.BackgroundColor.Indexed);
        }
        [Ignore]
        [TestMethod]
        public void ReadBug11()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\sample.xlsx"));
            var ws = package.Workbook.Worksheets[1];
            var pck2 = new ExcelPackage();
            pck2.Workbook.Worksheets.Add("Test", ws);
            pck2.SaveAs(new FileInfo(@"c:\temp\SampleNew.xlsx"));
        }
        [Ignore]
        [TestMethod]
        public void ReadConditionalFormatting()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\cf2.xlsx"));
            var ws = package.Workbook.Worksheets[1];
            ws.Cells["A1"].Value = 1;
            Assert.AreEqual(ws.ConditionalFormatting[6].Type, eExcelConditionalFormattingRuleType.Equal);
            package.SaveAs(new FileInfo(@"c:\temp\condFormTest.xlsx"));
        }
        [Ignore]
        [TestMethod]
        public void ReadStyleBug()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\acquisitions-1993-2.xlsx"));
            var ws = package.Workbook.Worksheets[1];
            ws.Cells["A1"].Value = 1;
            package.SaveAs(new FileInfo(@"c:\temp\condFormTest.xlsx"));
        }
        [Ignore]
        [TestMethod]
        public void ReadURL()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\url.xlsx"));
            var ws = package.Workbook.Worksheets[1];
            ws.Cells["A1"].Value = 1;
            package.SaveAs(new FileInfo(@"c:\temp\condFormTest.xlsx"));
        }
        [Ignore]
        [TestMethod]
        public void ReadNameError()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\names2.xlsx"));
            var ws = package.Workbook.Worksheets[1];
            ws.Cells["A1"].Value = 1;
            package.SaveAs(new FileInfo(@"c:\temp\TestTableSave.xlsx"));
        }
        [TestMethod]
        public void ReadBug12()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\test4.xlsx"));
            var ws = package.Workbook.Worksheets[1];
            ws.Cells["A1"].Value = 1;
            //ws.Column(0).Style.Font.Bold = true;
            package.SaveAs(new FileInfo(@"c:\temp\bug2.xlsx"));
        }
        [Ignore]
        [TestMethod]
        public void ReadBug13()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\original.xlsx"));
            var ws = package.Workbook.Worksheets[1];
            package.Workbook.Calculate(new OfficeOpenXml.FormulaParsing.ExcelCalculationOption() { AllowCirculareReferences = true });
            package.SaveAs(new FileInfo(@"c:\temp\bug2.xlsx"));
        }
        [Ignore]
        [TestMethod]
        public void ReadBug14()
        {
            var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Comment");
            ws.Cells["A1"].AddComment("Test av kommentar", "J");
            ws.Comments.RemoveAt(0);
            package.SaveAs(new FileInfo(@"c:\temp\bug\CommentTest.xlsx"));
        }
        [Ignore]
        [TestMethod]
        public void ReadBug15()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\ColumnMaxError.xlsx"));
            var ws = package.Workbook.Worksheets[1];
            var col = ws.Column(1);
            col.Style.Fill.PatternType = ExcelFillStyle.Solid;
            col.Style.Fill.BackgroundColor.SetColor(Color.Red);
            package.SaveAs(new FileInfo(@"c:\temp\bug2.xlsx"));
        }
        #region "Threading Cellstore Test"
        public int _threadCount=0;
        ExcelPackage _pckThread;
        [TestMethod, Ignore]
        public void ThreadingTest()
        {
            _pckThread = new ExcelPackage();
            var ws = _pckThread.Workbook.Worksheets.Add("Threading");

            for (int t = 0; t < 20; t++)
            {
                var ts=new ThreadState(Finnished)
                {
                    ws=ws,
                    StartRow=1+(t*1000),
                    Rows=1000,                    
                };
                var tstart=new ThreadStart(ts.StartLoad);                                
                var thread = new Thread(tstart);
                _threadCount++;
                thread.Start();
            }
            while (1 == 1)
            {                
                if (_threadCount == 0)
                {
                    _pckThread.SaveAs(new FileInfo("c:\\temp\\thread.xlsx"));
                    break;
                }
                Thread.Sleep(1000);
            }
        }
        public void Finnished()
        {
            _threadCount--;
        }
        private class ThreadState
        {
            public ThreadState(cbFinished cb)
            {
                _cb = cb;
            }
            public ExcelWorksheet ws { get; set; }
            public int StartRow { get; set; }
            public int Rows { get; set; }
            public delegate void cbFinished();
            public cbFinished _cb;
            public void StartLoad()
            {
                for(int row=StartRow;row<StartRow+Rows;row++)
                {
                    for (int col = 1; col < 100; col++)
                    {
                        ws.SetValue(row,col,string.Format("row {0} col {1}", row,col));
                    }
                }
                _cb();
            }
        }
        #endregion
        [Ignore]
        [TestMethod]
        public void TestInvalidVBA()
        {
            const string infile=@"C:\temp\bug\Infile.xlsm";
            const string outfile=@"C:\temp\bug\Outfile.xlsm";
            ExcelPackage ep;

            using (FileStream fs = File.OpenRead(infile))
            {
                ep = new ExcelPackage(fs);
            }

            using (FileStream fs = File.OpenWrite(outfile))
            {
                ep.SaveAs(fs);
            }

            using (FileStream fs = File.OpenRead(outfile))
            {
                ep = new ExcelPackage(fs);
            }

            using (FileStream fs = File.OpenWrite(outfile))
            {
                ep.SaveAs(fs);
            }            
        }
        [Ignore]
        [TestMethod]
        public void StreamTest()
        {
            using (var templateStream = File.OpenRead(@"c:\temp\thread.xlsx"))
            {

                using (var outStream = File.Open(@"c:\temp\streamOut.xlsx", FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                {
                    using (var package = new ExcelPackage(outStream, templateStream))
                    {
                        package.Workbook.Worksheets[1].Cells["A1"].Value = 1;
                        // Create more content
                        package.Save();
                    }
                }
            }
        }
        [TestMethod]
        public void test()
        { 
            CreateXlsxSheet(@"C:\temp\bug\test4.xlsx", 4, 4);
            CreateXlsxSheet(@"C:\temp\bug\test25.xlsx", 25, 25); 
        }
        [Ignore]
        [TestMethod]
        public void I15038()
        {
            using(var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\15038.xlsx")))
            {
                var ws=p.Workbook.Worksheets[1];
            
            }
        }
        [Ignore]
        [TestMethod]
        public void I15039()
        {
            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\15039.xlsm")))
            {
                var ws = p.Workbook.Worksheets[1];

                p.SaveAs(new FileInfo(@"c:\temp\bug\15039-saved.xlsm"));
            }
        }
        [Ignore]
        [TestMethod]
        public void I15030()
        {
            using (var newPack = new ExcelPackage(new FileInfo(@"c:\temp\bug\I15030.xlsx")))
            {
                var wkBk = newPack.Workbook.Worksheets[1];
                var cell = wkBk.Cells["A1"];
                if (cell.Comment != null)
                {
                    cell.Comment.Text = "Hello edited comments";
                }
                newPack.SaveAs(new FileInfo(@"c:\temp\bug\15030-save.xlsx"));
            }
        }
        [Ignore]
        [TestMethod]
        public void I15014()
        {
            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\ClassicWineCompany.xlsx")))
            {
                var ws = p.Workbook.Worksheets[1];
                Assert.AreEqual("SFFSectionHeading00", ws.Cells[5, 2].StyleName);
            }
        }
        [Ignore]
        [TestMethod]
        public void I15043()
        {
            using (var p = new ExcelPackage(new FileInfo(@"C:\temp\bug\EPPlusTest\EPPlusTest\EPPlusTest\example.xlsx")))
            {
                var ws = p.Workbook.Worksheets[1];
                p.Workbook.Worksheets.Copy(ws.Name, "Copy");
            }
        }
        [Ignore]
        [TestMethod]
        public void whitespace()
        {
            using (var p = new ExcelPackage(new FileInfo(@"C:\temp\book1.xlsx")))
            {
                var ws = p.Workbook.Worksheets[1];
                p.Workbook.Worksheets.Copy(ws.Name, "Copy");
            }
        }
        private static void CreateXlsxSheet(string pFileName, int pRows, int pColumns) 
        {
            if (File.Exists(pFileName)) File.Delete(pFileName);

            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(pFileName)))
            {
                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add("Testsheet");

                // Fill with data
                for (int row = 1; row <= pRows; row++)
                {
                    for (int column = 1; column <= pColumns; column++)
                    {
                        if (column > 1 && row > 2)
                        {
                            using (ExcelRange range = excelWorksheet.Cells[row, column])
                            {
                                range.Style.Numberformat.Format = "0";
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            }
                            excelWorksheet.Cells[row, column].Value = row * column;
                        }
                    }
                }

                // Try to style the first column, begining with row 3 which has no content yet...
                using (ExcelRange range = excelWorksheet.Cells[ExcelCellBase.GetAddress(3, 1, pRows, 1)])
                {
                    ExcelStyle style = range.Style;
                }

                // now I would add data to the first column (left out here)...
                excelPackage.Save();
            } 
        }    
    }
}
