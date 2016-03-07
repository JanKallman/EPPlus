using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestClass]
    public class CountIfsTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add("testsheet");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void ShouldHandleSingleNumericCriteria()
        {
            _worksheet.Cells["A1"].Value = 1;
            _worksheet.Cells["A2"].Value = 1;
            _worksheet.Cells["A3"].Value = 2;
            _worksheet.Cells["A4"].Formula = "COUNTIFS(A1:A3, 1)";
            _worksheet.Calculate();
            Assert.AreEqual(2d, _worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void ShouldHandleSingleRangeCriteria()
        {
            _worksheet.Cells["A1"].Value = 1;
            _worksheet.Cells["A2"].Value = 1;
            _worksheet.Cells["A3"].Value = 2;
            _worksheet.Cells["B1"].Value = 1;
            _worksheet.Cells["A4"].Formula = "COUNTIFS(A1:A3, B1)";
            _worksheet.Calculate();
            Assert.AreEqual(2d, _worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void ShouldHandleSingleNumericWildcardCriteria()
        {
            _worksheet.Cells["A1"].Value = 1;
            _worksheet.Cells["A2"].Value = 2;
            _worksheet.Cells["A3"].Value = 3;
            _worksheet.Cells["A4"].Formula = "COUNTIFS(A1:A3, \"<3\")";
            _worksheet.Calculate();
            Assert.AreEqual(2d, _worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void ShouldHandleSingleStringCriteria()
        {
            _worksheet.Cells["A1"].Value = "abc";
            _worksheet.Cells["A2"].Value = "def";
            _worksheet.Cells["A3"].Value = "def";
            _worksheet.Cells["A4"].Formula = "COUNTIFS(A1:A3, \"def\")";
            _worksheet.Calculate();
            Assert.AreEqual(2d, _worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void ShouldHandleSingleStringWildcardCriteria()
        {
            _worksheet.Cells["A1"].Value = "abc";
            _worksheet.Cells["A2"].Value = "def";
            _worksheet.Cells["A3"].Value = "def";
            _worksheet.Cells["A4"].Formula = "COUNTIFS(A1:A3, \"d*f\")";
            _worksheet.Calculate();
            Assert.AreEqual(2d, _worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void ShouldHandleNullRangeCriteria()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = 1;
            _worksheet.Cells["A3"].Value = null;
            _worksheet.Cells["A4"].Formula = "COUNTIFS(A1:A3, B1)";
            _worksheet.Calculate();
            Assert.AreEqual(0d, _worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void ShouldHandleMultipleRangesAndCriterias()
        {
            _worksheet.Cells["A1"].Value = "abc";
            _worksheet.Cells["A2"].Value = "def";
            _worksheet.Cells["A3"].Value = "def";
            _worksheet.Cells["A4"].Value = "def";
            _worksheet.Cells["B1"].Value = 1;
            _worksheet.Cells["B2"].Value = 2;
            _worksheet.Cells["B3"].Value = 3;
            _worksheet.Cells["B4"].Value = 2;
            _worksheet.Cells["C1"].Value = null;
            _worksheet.Cells["C2"].Value = 200;
            _worksheet.Cells["C3"].Value = 3;
            _worksheet.Cells["C4"].Value = 2;
            _worksheet.Cells["A5"].Formula = "COUNTIFS(A1:A4, \"d*f\", B1:B4; 2; C1:C4; 200)";
            _worksheet.Calculate();
            Assert.AreEqual(1d, _worksheet.Cells["A5"].Value);
        }

        [TestMethod]
        public void ShouldSetErrorValueIfNumberOfCellsInRangesDiffer()
        {
            _worksheet.Cells["A1"].Value = "abc";
            _worksheet.Cells["A2"].Value = "def";
            _worksheet.Cells["A3"].Value = "def";
            _worksheet.Cells["A4"].Value = "def";
            _worksheet.Cells["B1"].Value = 1;
            _worksheet.Cells["B2"].Value = 2;
            _worksheet.Cells["B3"].Value = 3;
            _worksheet.Cells["B4"].Value = 2;
            _worksheet.Cells["A5"].Formula = "COUNTIFS(A1:A4, \"d*f\", B1:B3; 2)";
            _worksheet.Calculate();
             Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), _worksheet.Cells["A5"].Value);
        }
    }
}
