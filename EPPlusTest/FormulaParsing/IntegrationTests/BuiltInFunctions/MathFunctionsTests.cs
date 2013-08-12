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
    public class MathFunctionsTests : FormulaParserTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            var excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(excelDataProvider);
        }

        [TestMethod]
        public void PowerShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Power(3, 3)");
            Assert.AreEqual(27d, result);
        }

        [TestMethod]
        public void SqrtShouldReturnCorrectResult()
        {
            var result = _parser.Parse("sqrt(9)");
            Assert.AreEqual(3d, result);
        }

        [TestMethod]
        public void PiShouldReturnCorrectResult()
        {
            var expectedValue = (double)Math.Round(Math.PI, 14);
            var result = _parser.Parse("Pi()");
            Assert.AreEqual(expectedValue, result);
        }

        [TestMethod]
        public void CeilingShouldReturnCorrectResult()
        {
            var expectedValue = 22.4d;
            var result = _parser.Parse("ceiling(22.35, 0.1)");
            Assert.AreEqual(expectedValue, result);
        }

        [TestMethod]
        public void FloorShouldReturnCorrectResult()
        {
            var expectedValue = 22.3d;
            var result = _parser.Parse("Floor(22.35, 0.1)");
            Assert.AreEqual(expectedValue, result);
        }

        [TestMethod]
        public void SumShouldReturnCorrectResultWithInts()
        {
            var result = _parser.Parse("sum(1, 2)");
            Assert.AreEqual(3d, result);
        }

        [TestMethod]
        public void SumShouldReturnCorrectResultWithDecimals()
        {
            var result = _parser.Parse("sum(1,2.5)");
            Assert.AreEqual(3.5d, result);
        }

        [TestMethod]
        public void SumShouldReturnCorrectResultWithEnumerable()
        {
            var result = _parser.Parse("sum({1,2,3}, 2.5)");
            Assert.AreEqual(8.5d, result);
        }

        [TestMethod]
        public void SumIfShouldReturnCorrectResult()
        {
            var result = _parser.Parse("sumIf({1,2,3,2}, 2)");
            Assert.AreEqual(4d, result);
        }

        [TestMethod]
        public void StdevShouldReturnAResult()
        {
            var result = _parser.Parse("stdev(1,2,3,4)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void StdevPShouldReturnAResult()
        {
            var result = _parser.Parse("stdevp(2,3,4)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void ExpShouldReturnAResult()
        {
            var result = _parser.Parse("exp(4)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void MaxShouldReturnAResult()
        {
            var result = _parser.Parse("Max(4, 5)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void MaxaShouldReturnAResult()
        {
            var result = _parser.Parse("Maxa(4, 5)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void MinShouldReturnAResult()
        {
            var result = _parser.Parse("min(4, 5)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void AverageShouldReturnAResult()
        {
            var result = _parser.Parse("Average(2, 2, 2)");
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void RoundShouldReturnAResult()
        {
            var result = _parser.Parse("Round(2.2, 0)");
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void SqrtPiShouldReturnAResult()
        {
            var result = _parser.Parse("SqrtPi(2.2)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void IntShouldReturnAResult()
        {
            var result = _parser.Parse("Int(2.9)");
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void RandShouldReturnAResult()
        {
            var result = _parser.Parse("Rand()");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void RandBetweenShouldReturnAResult()
        {
            var result = _parser.Parse("RandBetween(1,2)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void CountShouldReturnAResult()
        {
            var result = _parser.Parse("Count(1,2,2,'4')");
            Assert.AreEqual(3d, result);
        }

        [TestMethod]
        public void CountAShouldReturnAResult()
        {
            var result = _parser.Parse("CountA(1,2,2,'', 'a')");
            Assert.AreEqual(4d, result);
        }

        [TestMethod]
        public void VarShouldReturnAResult()
        {
            var result = _parser.Parse("Var(1,2,3)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void VarPShouldReturnAResult()
        {
            var result = _parser.Parse("VarP(1,2,3)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void ModShouldReturnAResult()
        {
            var result = _parser.Parse("Mod(5,2)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void SubtotalShouldReturnAResult()
        {
            var result = _parser.Parse("Subtotal(1, 10, 20)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void ProductShouldReturnAResult()
        {
            var result = _parser.Parse("Product(1,2,3)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void CosShouldReturnAResult()
        {
            var result = _parser.Parse("Cos(2)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void CoshShouldReturnAResult()
        {
            var result = _parser.Parse("Cosh(2)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void SinShouldReturnAResult()
        {
            var result = _parser.Parse("Sin(2)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void SinhShouldReturnAResult()
        {
            var result = _parser.Parse("Sinh(2)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void TanShouldReturnAResult()
        {
            var result = _parser.Parse("Tan(2)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void AtanShouldReturnAResult()
        {
            var result = _parser.Parse("Atan(2)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void Atan2ShouldReturnAResult()
        {
            var result = _parser.Parse("Atan2(2,1)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void TanhShouldReturnAResult()
        {
            var result = _parser.Parse("Tanh(2)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void LogShouldReturnAResult()
        {
            var result = _parser.Parse("Log(2, 2)");
            Assert.AreEqual(1d, result);
        }

        [TestMethod]
        public void Log10ShouldReturnAResult()
        {
            var result = _parser.Parse("Log10(2)");
            Assert.IsInstanceOfType(result, typeof(double));
        }
    }
}
