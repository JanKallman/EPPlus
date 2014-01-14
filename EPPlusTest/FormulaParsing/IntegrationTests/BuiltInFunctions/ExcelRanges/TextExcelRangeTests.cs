using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Calculation;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions.ExcelRanges
{
    [TestClass]
    public class TextExcelRangeTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add("Test");

            _worksheet.Cells["A1"].Value = 1;
            _worksheet.Cells["A2"].Value = 3;
            _worksheet.Cells["A3"].Value = 6;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void ExactShouldReturnTrueWhenEqualValues()
        {
            _worksheet.Cells["A2"].Value = 1d;
            _worksheet.Cells["A4"].Formula = "EXACT(A1,A2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void FindShouldReturnIndex()
        {
            _worksheet.Cells["A1"].Value = "hopp";
            _worksheet.Cells["A2"].Value = "hej hopp";
            _worksheet.Cells["A4"].Formula = "Find(A1,A2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(5, result);
        }
    }
}
