using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class SubtotalTests : FormulaParserTestBase
    {
        private ExcelDataProvider _excelDataProvider;

        [TestInitialize]
        public void Setup()
        {
            _excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(_excelDataProvider);
        }

        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren()
        {
            _excelDataProvider
                .Stub(x => x.GetRangeFormula(string.Empty, 0, 0))
                .Return("SUBTOTAL(9, A2:A3)");
            _excelDataProvider
                .Stub(x => x.GetRangeValues("A2:A3"))
                .Return(new List<object> { "SUBTOTAL(9, A5:A6)", 2d});
            _excelDataProvider
                .Stub(x => x.GetRangeValues("A5:A6"))
                .Return(new List<object> { 2d, 2d });
            var result = _parser.ParseAt("A1");
            Assert.AreEqual(2d, result);
        }
    }
}
