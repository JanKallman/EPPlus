using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.IntegrationTests
{
    [TestClass]
    public class OperatorsTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _ws;
        private readonly ExcelErrorValue DivByZero = ExcelErrorValue.Create(eErrorType.Div0);

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _ws = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void DivByZeroShouldReturnError()
        {
            var result = _ws.Calculate("10/0 + 3");
            Assert.AreEqual(DivByZero, result);
        }
    }
}
