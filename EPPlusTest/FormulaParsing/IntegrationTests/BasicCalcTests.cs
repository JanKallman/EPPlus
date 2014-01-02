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
    public class BasicCalcTests : FormulaParserTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            var excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(excelDataProvider);
        }

        [TestMethod]
        public void ShouldAddIntegersCorrectly()
        {
            var result = _parser.Parse("1 + 2");
            Assert.AreEqual(3d, result);
        }

        [TestMethod]
        public void ShouldSubtractIntegersCorrectly()
        {
            var result = _parser.Parse("2 - 1");
            Assert.AreEqual(1d, result);
        }

        [TestMethod]
        public void ShouldMultiplyIntegersCorrectly()
        {
            var result = _parser.Parse("2 * 3");
            Assert.AreEqual(6d, result);
        }

        [TestMethod]
        public void ShouldDivideIntegersCorrectly()
        {
            var result = _parser.Parse("8 / 4");
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void ShouldDivideDecimalWithIntegerCorrectly()
        {
            var result = _parser.Parse("2.5/2");
            Assert.AreEqual(1.25d, result);
        }

        [TestMethod]
        public void ShouldHandleExpCorrectly()
        {
            var result = _parser.Parse("2 ^ 4");
            Assert.AreEqual(16d, result);
        }

        [TestMethod]
        public void ShouldHandleExpWithDecimalCorrectly()
        {
            var result = _parser.Parse("2.5 ^ 2");
            Assert.AreEqual(6.25d, result);
        }

        [TestMethod]
        public void ShouldMultiplyDecimalWithDecimalCorrectly()
        {
            var result = _parser.Parse("2.5 * 1.5");
            Assert.AreEqual(3.75d, result);
        }

        [TestMethod]
        public void ThreeGreaterThanTwoShouldBeTrue()
        {
            var result = _parser.Parse("3 > 2");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void ThreeLessThanTwoShouldBeFalse()
        {
            var result = _parser.Parse("3 < 2");
            Assert.IsFalse((bool)result);
        }

        [TestMethod]
        public void ThreeLessThanOrEqualToThreeShouldBeTrue()
        {
            var result = _parser.Parse("3 <= 3");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void ThreeLessThanOrEqualToTwoDotThreeShouldBeFalse()
        {
            var result = _parser.Parse("3 <= 2.3");
            Assert.IsFalse((bool)result);
        }

        [TestMethod]
        public void ThreeGreaterThanOrEqualToThreeShouldBeTrue()
        {
            var result = _parser.Parse("3 >= 3");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void TwoDotTwoGreaterThanOrEqualToThreeShouldBeFalse()
        {
            var result = _parser.Parse("2.2 >= 3");
            Assert.IsFalse((bool)result);
        }

        [TestMethod]
        public void TwelveAndTwelveShouldBeEqual()
        {
            var result = _parser.Parse("2=2");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void TenPercentShouldBe0Point1()
        {
            var result = _parser.Parse("10%");
            Assert.AreEqual(0.1, result);
        }

        [TestMethod]
        public void ShouldHandleMultiplePercentSigns()
        {
            var result = _parser.Parse("10%%");
            Assert.AreEqual(0.001, result);
        }

        [TestMethod]
        public void ShouldHandlePercentageOnFunctionResult()
        {
            var result = _parser.Parse("SUM(1;2;3)%");
            Assert.AreEqual(0.06, result);
        }
    }
}
