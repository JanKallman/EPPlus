using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions.ExcelRanges
{
    [TestClass]
    public class WorksheetRefsTest
    {
        private ExcelPackage _package;
        private ExcelWorksheet _firstSheet;
        private ExcelWorksheet _secondSheet;

        [TestInitialize]
        public void Init()
        {
            _package = new ExcelPackage();
            _firstSheet = _package.Workbook.Worksheets.Add("sheet1");
            _secondSheet = _package.Workbook.Worksheets.Add("sheet2");
            _firstSheet.Cells["A1"].Value = 1;
            _firstSheet.Cells["A2"].Value = 2;
        }

        [TestCleanup]
        public void Cleanup()
        {
            
            _package.Dispose();
        }

        [TestMethod]
        public void ShouldHandleReferenceToOtherSheet()
        {
            _secondSheet.Cells["A1"].Formula = "SUM('sheet1'!A1:A2)";
            _secondSheet.Calculate();
            Assert.AreEqual(3d, _secondSheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldHandleReferenceToOtherSheetWithComplexName()
        {
            var sheet = _package.Workbook.Worksheets.Add("ab#k..2");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            _secondSheet.Cells["A1"].Formula = "SUM('ab#k..2'A1:A2)";
            _secondSheet.Calculate();
            Assert.AreEqual(3d, _secondSheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void ShouldHandleInvalidRef()
        {
            var sheet = _package.Workbook.Worksheets.Add("ab#k..2");
            sheet.Cells["A1"].Value = 1;
            sheet.Cells["A2"].Value = 2;
            _secondSheet.Cells["A1"].Formula = "SUM('ab#k..2A1:A2')";
            _secondSheet.Calculate();
            Assert.IsInstanceOfType(_secondSheet.Cells["A1"].Value, typeof(ExcelErrorValue));
        }
    }
}
