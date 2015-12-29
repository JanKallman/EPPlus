using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using Rhino.Mocks;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class InformationFunctionsTests : FormulaParserTestBase
    {
        private ExcelDataProvider _excelDataProvider;

        [TestInitialize]
        public void Setup()
        {
            _excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(_excelDataProvider);
        }

        [TestMethod]
        public void IsBlankShouldReturnCorrectValue()
        {
            var result = _parser.Parse("ISBLANK(A1)");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void IsNumberShouldReturnCorrectValue()
        {
            var result = _parser.Parse("ISNUMBER(10/2)");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void IsErrorShouldReturnTrueWhenDivBy0()
        {
            var result = _parser.Parse("ISERROR(10/0)");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void IsTextShouldReturnTrueWhenReferencedCellContainsText()
        {
            using(var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "Abc";
                sheet.Cells["A2"].Formula = "ISTEXT(A1)";
                sheet.Calculate();
                var result = sheet.Cells["A2"].Value;
                Assert.IsTrue((bool)result);
            }
        }

        [TestMethod]
        public void IsErrShouldReturnFalseIfErrorCodeIsNa()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = ExcelErrorValue.Parse("#N/A");
                sheet.Cells["A2"].Formula = "ISERR(A1)";
                sheet.Calculate();
                var result = sheet.Cells["A2"].Value;
                Assert.IsFalse((bool)result);
            }
        }

        [TestMethod]
        public void IsNaShouldReturnTrueCodeIsNa()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = ExcelErrorValue.Parse("#N/A");
                sheet.Cells["A2"].Formula = "ISNA(A1)";
                sheet.Calculate();
                var result = sheet.Cells["A2"].Value;
                Assert.IsTrue((bool)result);
            }
        }

        [TestMethod]
        public void ErrorTypeShouldReturnCorrectErrorCodes()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = ExcelErrorValue.Create(eErrorType.Null);
                sheet.Cells["B1"].Formula = "ERROR.TYPE(A1)";
                sheet.Cells["A2"].Value = ExcelErrorValue.Create(eErrorType.Div0);
                sheet.Cells["B2"].Formula = "ERROR.TYPE(A2)";
                sheet.Cells["A3"].Value = ExcelErrorValue.Create(eErrorType.Value);
                sheet.Cells["B3"].Formula = "ERROR.TYPE(A3)";
                sheet.Cells["A4"].Value = ExcelErrorValue.Create(eErrorType.Ref);
                sheet.Cells["B4"].Formula = "ERROR.TYPE(A4)";
                sheet.Cells["A5"].Value = ExcelErrorValue.Create(eErrorType.Name);
                sheet.Cells["B5"].Formula = "ERROR.TYPE(A5)";
                sheet.Cells["A6"].Value = ExcelErrorValue.Create(eErrorType.Num);
                sheet.Cells["B6"].Formula = "ERROR.TYPE(A6)";
                sheet.Cells["A7"].Value = ExcelErrorValue.Create(eErrorType.NA);
                sheet.Cells["B7"].Formula = "ERROR.TYPE(A7)";
                sheet.Cells["A8"].Value = 10;
                sheet.Cells["B8"].Formula = "ERROR.TYPE(A8)";
                sheet.Calculate();
                var nullResult = sheet.Cells["B1"].Value;
                var div0Result = sheet.Cells["B2"].Value;
                var valueResult = sheet.Cells["B3"].Value;
                var refResult = sheet.Cells["B4"].Value;
                var nameResult = sheet.Cells["B5"].Value;
                var numResult = sheet.Cells["B6"].Value;
                var naResult = sheet.Cells["B7"].Value;
                var noErrorResult = sheet.Cells["B8"].Value;
                Assert.AreEqual(1, nullResult, "Null error was not 1");
                Assert.AreEqual(2, div0Result, "Div0 error was not 2");
                Assert.AreEqual(3, valueResult, "Value error was not 3");
                Assert.AreEqual(4, refResult, "Ref error was not 4");
                Assert.AreEqual(5, nameResult, "Name error was not 5");
                Assert.AreEqual(6, numResult, "Num error was not 6");
                Assert.AreEqual(7, naResult, "NA error was not 7");
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), noErrorResult, "No error did not return N/A error");
            }
        }
    }
}
