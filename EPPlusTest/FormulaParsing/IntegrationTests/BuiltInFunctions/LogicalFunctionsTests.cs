using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using Rhino.Mocks;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class LogicalFunctionsTests : FormulaParserTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            var excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(excelDataProvider);
        }

        [TestMethod]
        public void IfShouldReturnCorrectResult()
        {
            var result = _parser.Parse("If(2 < 3, 1, 2)");
            Assert.AreEqual(1d, result);
        }

        [TestMethod]
        public void IIfShouldReturnCorrectResultWhenInnerFunctionExists()
        {
            var result = _parser.Parse("If(NOT(Or(true, FALSE)), 1, 2)");
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void NotShouldReturnCorrectResult()
        {
            var result = _parser.Parse("not(true)");
            Assert.IsFalse((bool)result);

            result = _parser.Parse("NOT(false)");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void AndShouldReturnCorrectResult()
        {
            var result = _parser.Parse("And(true, 1)");
            Assert.IsTrue((bool)result);

            result = _parser.Parse("AND(true, true, 1, false)");
            Assert.IsFalse((bool)result);
        }

        [TestMethod]
        public void OrShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Or(FALSE, 0)");
            Assert.IsFalse((bool)result);

            result = _parser.Parse("OR(true, true, 1, false)");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void TrueShouldReturnCorrectResult()
        {
            var result = _parser.Parse("True()");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void FalseShouldReturnCorrectResult()
        {
            var result = _parser.Parse("False()");
            Assert.IsFalse((bool)result);
        }
    }
}
