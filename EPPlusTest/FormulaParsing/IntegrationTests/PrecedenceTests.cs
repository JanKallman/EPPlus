using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using Rhino.Mocks;

namespace EPPlusTest.FormulaParsing.IntegrationTests
{
    [TestClass]
    public class PrecedenceTests : FormulaParserTestBase
    {

        [TestInitialize]
        public void Setup()
        {
            var excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(excelDataProvider);
        }

        [TestMethod]
        public void ShouldCaluclateUsingPrecedenceMultiplyBeforeAdd()
        {
            var result = _parser.Parse("4 + 6 * 2");
            Assert.AreEqual(16d, result);
        }

        [TestMethod]
        public void ShouldCaluclateUsingPrecedenceDivideBeforeAdd()
        {
            var result = _parser.Parse("4 + 6 / 2");
            Assert.AreEqual(7d, result);
        }

        [TestMethod]
        public void ShouldCalculateTwoGroupsUsingDivideAndMultiplyBeforeSubtract()
        {
            var result = _parser.Parse("4/2 + 3 * 3");
            Assert.AreEqual(11d, result);
        }

        [TestMethod]
        public void ShouldCalculateExpressionWithinParenthesisBeforeMultiply()
        {
            var result = _parser.Parse("(2+4) * 2");
            Assert.AreEqual(12d, result);
        }

        [TestMethod]
        public void ShouldConcatAfterAdd()
        {
            var result = _parser.Parse("2 + 4 & \"abc\"");
            Assert.AreEqual("6abc", result);
        }

        [TestMethod]
        public void Bugfixtest()
        {
            var result = _parser.Parse("(1+2)+3^2");
        }
    }
}
