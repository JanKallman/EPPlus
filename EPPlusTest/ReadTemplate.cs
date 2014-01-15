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
        [Ignore]
        [TestMethod]
        public void ReadBug12()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\grid.xlsx"));
            var ws = package.Workbook.Worksheets[1];
            ws.Cells["A1"].Value = 1;
            package.SaveAs(new FileInfo(@"c:\temp\bug2.xlsx"));
        }
    }
}
