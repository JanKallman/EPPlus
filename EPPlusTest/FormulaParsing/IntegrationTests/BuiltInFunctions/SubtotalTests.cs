using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml;
using OfficeOpenXml.Calculation;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class SubtotalTests : FormulaParserTestBase
    {
        private ExcelWorksheet _worksheet;
        private ExcelPackage _package;

        [TestInitialize]
        public void Setup()
        {
            _package = new ExcelPackage(new MemoryStream());
            _worksheet = _package.Workbook.Worksheets.Add("Test");
            _parser = _worksheet.Workbook.FormulaParser;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Sum()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(9, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Avg()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(1, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Count()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(2, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(1d, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_CountA()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(3, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(1d, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Max()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(4, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Min()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(5, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Product()
        {
            _worksheet.Cells["A1"].Formula = "SUBTOTAL(6, A2:A3)";
            _worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
            _worksheet.Cells["A3"].Value = 2d;
            _worksheet.Cells["A5"].Value = 2d;
            _worksheet.Calculate();
            var result = _worksheet.Cells["A1"].Value;
            Assert.AreEqual(2d, result);
        }
    }
}
