using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class RefAndLookupTests : FormulaParserTestBase
    {
        private ExcelDataProvider _excelDataProvider;
        const string WorksheetName = null;
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add("Test");
            _excelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => _excelDataProvider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(10, 1));
            A.CallTo(() => _excelDataProvider.GetWorkbookNameValues()).Returns(new ExcelNamedRangeCollection(_package.Workbook));
            _parser = new FormulaParser(_excelDataProvider);    
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void VLookupShouldReturnCorrespondingValue()
        {
            using(var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("test");
                var lookupAddress = "A1:B2";
                ws.Cells["A1"].Value = 1;
                ws.Cells["B1"].Value = 1;
                ws.Cells["A2"].Value = 2;
                ws.Cells["B2"].Value = 5;
                ws.Cells["A3"].Formula = "VLOOKUP(2, " + lookupAddress + ", 2)";
                ws.Calculate();
                var result = ws.Cells["A3"].Value;
                Assert.AreEqual(5, result);
            }
        }

        [TestMethod]
        public void VLookupShouldReturnClosestValueBelowIfLastArgIsTrue()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("test");
                var lookupAddress = "A1:B2";
                ws.Cells["A1"].Value = 3;
                ws.Cells["B1"].Value = 1;
                ws.Cells["A2"].Value = 5;
                ws.Cells["B2"].Value = 5;
                ws.Cells["A3"].Formula = "VLOOKUP(4, " + lookupAddress + ", 2, true)";
                ws.Calculate();
                var result = ws.Cells["A3"].Value;
                Assert.AreEqual(1, result);
            }
        }

        [TestMethod]
        public void HLookupShouldReturnCorrespondingValue()
        {
            var lookupAddress = "A1:B2";
            _worksheet.Cells["A1"].Value = 1;
            _worksheet.Cells["B1"].Value = 2;
            _worksheet.Cells["A2"].Value = 2;
            _worksheet.Cells["B2"].Value = 5;
            _worksheet.Cells["A3"].Formula = "HLOOKUP(2, " + lookupAddress + ", 2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A3"].Value;
            Assert.AreEqual(5, result);
        }

        [TestMethod]
        public void HLookupShouldReturnClosestValueBelowIfLastArgIsTrue()
        {
            var lookupAddress = "A1:B2";
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells[1, 1].Value = 3;
                s.Cells[1, 2].Value = 5;
                s.Cells[2, 1].Value = 1;
                s.Cells[2, 2].Value = 2;
                s.Cells[5, 5].Formula = "HLOOKUP(4, " + lookupAddress + ", 2, true)";
                s.Calculate();
                Assert.AreEqual(1, s.Cells[5,5].Value);
            }
        }

        [TestMethod]
        public void LookupShouldReturnMatchingValue()
        {
            var lookupAddress = "A1:B2";
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells[1, 1].Value = 3;
                s.Cells[1, 2].Value = 5;
                s.Cells[2, 1].Value = 4;
                s.Cells[2, 2].Value = 1;
                s.Cells[5, 5].Formula = "LOOKUP(4, " + lookupAddress + ")";
                s.Calculate();
                Assert.AreEqual(1, s.Cells[5, 5].Value);
            }
            //    A.CallTo(() => _excelDataProvider.GetCellValue(WorksheetName,1, 1)).Returns(3);
            //A.CallTo(() => _excelDataProvider.GetCellValue(WorksheetName,1, 2)).Returns(5);
            //A.CallTo(() => _excelDataProvider.GetCellValue(WorksheetName,2, 1)).Returns(4);
            //A.CallTo(() => _excelDataProvider.GetCellValue(WorksheetName,2, 2)).Returns(1);
            //var result = _parser.Parse("LOOKUP(4, " + lookupAddress + ")");
            //Assert.AreEqual(1, result);
        }
           
        [TestMethod]
        public void MatchShouldReturnIndexOfMatchingValue()
        {
            var lookupAddress = "A1:A2";
            A.CallTo(() => _excelDataProvider.GetCellValue(WorksheetName,1, 1)).Returns(3);
            A.CallTo(() => _excelDataProvider.GetCellValue(WorksheetName,1, 2)).Returns(5);
            var result = _parser.Parse("MATCH(3, " + lookupAddress + ")");
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void RowShouldReturnRowNumber()
        {
            A.CallTo(() => _excelDataProvider.GetRangeFormula("", 4, 1)).Returns("Row()");
            var result = _parser.ParseAt("A4");
            Assert.AreEqual(4, result);
        }

        [TestMethod]
        public void RowSholdHandleReference()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "ROW(A4)";
                s1.Calculate();
                Assert.AreEqual(4, s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void ColumnShouldReturnRowNumber()
        {
            A.CallTo(() => _excelDataProvider.GetRangeFormula("", 4, 2)).Returns("Column()");
            var result = _parser.ParseAt("B4");
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void ColumnSholdHandleReference()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "COLUMN(B4)";
                s1.Calculate();
                Assert.AreEqual(2, s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void RowsShouldReturnNbrOfRows()
        {
            A.CallTo(() => _excelDataProvider.GetRangeFormula("", 4, 1)).Returns("Rows(A5:B7)");
            var result = _parser.ParseAt("A4");
            Assert.AreEqual(3, result);
        }

        [TestMethod]
        public void ColumnsShouldReturnNbrOfCols()
        {
            A.CallTo(() => _excelDataProvider.GetRangeFormula("", 4, 1)).Returns("Columns(A5:B7)");
            var result = _parser.ParseAt("A4");
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void ChooseShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Choose(1, \"A\", \"B\")");
            Assert.AreEqual("A", result);
        }

        [TestMethod]
        public void AddressShouldReturnCorrectResult()
        {
            A.CallTo(() => _excelDataProvider.ExcelMaxRows).Returns(12345);
            var result = _parser.Parse("Address(1, 1)");
            Assert.AreEqual("$A$1", result);
        }

        [TestMethod]
        public void IndirectShouldReturnARange()
        {
            using (var package = new ExcelPackage(new MemoryStream()))
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["A1:A2"].Value = 2;
                s1.Cells["A3"].Formula = "SUM(Indirect(\"A1:A2\"))";
                s1.Calculate();
                Assert.AreEqual(4d, s1.Cells["A3"].Value);

                s1.Cells["A4"].Formula = "SUM(Indirect(\"A1:A\" & \"2\"))";
                s1.Calculate();
                Assert.AreEqual(4d, s1.Cells["A4"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldReturnASingleValue()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B3"].Value = 1d;
                s1.Cells["A5"].Formula = "OFFSET(A1, 2, 1)";
                s1.Calculate();
                Assert.AreEqual(1d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldReturnARange()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B1"].Value = 1d;
                s1.Cells["B2"].Value = 1d;
                s1.Cells["B3"].Value = 1d;
                s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1))";
                s1.Calculate();
                Assert.AreEqual(3d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldReturnARange2()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B1"].Value = 10d;
                s1.Cells["B2"].Value = 10d;
                s1.Cells["B3"].Value = 10d;
                s1.Cells["A5"].Formula = "COUNTA(OFFSET(Test!B1, 0, 0, Test!B2, Test!B3))";
                s1.Calculate();
                Assert.AreEqual(3d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void OffsetDirectReferenceToMultiRangeShouldSetValueError()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B1"].Value = 1d;
                s1.Cells["B2"].Value = 1d;
                s1.Cells["B3"].Value = 1d;
                s1.Cells["A5"].Formula = "OFFSET(A1:A3, 0, 1)";
                s1.Calculate();
                var result = s1.Cells["A5"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result);
            }
        }

        [TestMethod]
        public void OffsetShouldReturnARangeAccordingToWidth()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B1"].Value = 1d;
                s1.Cells["B2"].Value = 1d;
                s1.Cells["B3"].Value = 1d;
                s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1, 2))";
                s1.Calculate();
                Assert.AreEqual(2d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldReturnARangeAccordingToHeight()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["B1"].Value = 1d;
                s1.Cells["B2"].Value = 1d;
                s1.Cells["B3"].Value = 1d;
                s1.Cells["C1"].Value = 2d;
                s1.Cells["C2"].Value = 2d;
                s1.Cells["C3"].Value = 2d;
                s1.Cells["A5"].Formula = "SUM(OFFSET(A1:A3, 0, 1, 2, 2))";
                s1.Calculate();
                Assert.AreEqual(6d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void OffsetShouldCoverMultipleColumns()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("Test");
                s1.Cells["C1"].Value = 1d;
                s1.Cells["C2"].Value = 1d;
                s1.Cells["C3"].Value = 1d;
                s1.Cells["D1"].Value = 2d;
                s1.Cells["D2"].Value = 2d;
                s1.Cells["D3"].Value = 2d;
                s1.Cells["A5"].Formula = "SUM(OFFSET(A1:B3, 0, 2))";
                s1.Calculate();
                Assert.AreEqual(9d, s1.Cells["A5"].Value);
            }
        }

        [TestMethod, Ignore]
        public void VLookupShouldHandleNames()
        {
            using (var package = new ExcelPackage(new FileInfo(@"c:\temp\Book3.xlsx")))
            {
                var s1 = package.Workbook.Worksheets.First();
                var v = s1.Cells["X10"].Formula;
                //s1.Calculate();
                v = s1.Cells["X10"].Formula;
            }
        }

        [TestMethod]
        public void LookupShouldReturnFromResultVector()
        {
            var lookupAddress = "A1:A5";
            var resultAddress = "B1:B5";
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                //lookup_vector
                s.Cells[1, 1].Value = 4.14;
                s.Cells[2, 1].Value = 4.19;
                s.Cells[3, 1].Value = 5.17;
                s.Cells[4, 1].Value = 5.77;
                s.Cells[5, 1].Value = 6.39;
                //result_vector
                s.Cells[1, 2].Value = "red";
                s.Cells[2, 2].Value = "orange";
                s.Cells[3, 2].Value = "yellow";
                s.Cells[4, 2].Value = "green";
                s.Cells[5, 2].Value = "blue";
                //lookup_value
                s.Cells[1, 3].Value = 4.14;
                s.Cells[5, 5].Formula = "LOOKUP(C1, " + lookupAddress + ", " + resultAddress + ")";
                s.Calculate();
                Assert.AreEqual("red", s.Cells[5, 5].Value);
            }
        }
    }
}
