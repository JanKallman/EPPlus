using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Calculation;
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
    }
}
