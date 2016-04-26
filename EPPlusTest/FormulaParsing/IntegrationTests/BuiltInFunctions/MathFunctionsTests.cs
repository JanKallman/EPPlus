using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using Rhino.Mocks;
using OfficeOpenXml;

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
            var result = _parser.Parse("sum({1;2;3;-1}, 2.5)");
            Assert.AreEqual(7.5d, result);
        }

        [TestMethod]
        public void SumsqShouldReturnCorrectResultWithEnumerable()
        {
            var result = _parser.Parse("sumsq({2;3})");
            Assert.AreEqual(13d, result);
        }

        [TestMethod]
        public void SumIfShouldReturnCorrectResult()
        {
            var result = _parser.Parse("sumIf({1;2;3;2}, 2)");
            Assert.AreEqual(4d, result);
        }

        [TestMethod]
        public void SubtotalShouldNegateExpression()
        {
            var result = _parser.Parse("-subtotal(2;{1;2})");
            Assert.AreEqual(-2d, result);
        }

        [TestMethod]
        public void StdevShouldReturnAResult()
        {
            var result = _parser.Parse("stdev(1;2;3;4)");
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
        public void MinaShouldCalculateStringAs0()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B2"].Value = "a";
                sheet.Cells["A5"].Formula = "MINA(A1:B4)";
                sheet.Calculate();
                Assert.AreEqual(0d, sheet.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void AverageShouldReturnAResult()
        {
            var result = _parser.Parse("Average(2, 2, 2)");
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void AverageShouldReturnDiv0IfEmptyCell()
        {
            using(var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("test");
                ws.Cells["A2"].Formula = "AVERAGE(A1)";
                ws.Calculate();
                Assert.AreEqual("#DIV/0!", ws.Cells["A2"].Value.ToString());
            }
        }

        [TestMethod]
        public void RoundShouldReturnAResult()
        {
            var result = _parser.Parse("Round(2.2, 0)");
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void RounddownShouldReturnAResult()
        {
            var result = _parser.Parse("Rounddown(2.99, 1)");
            Assert.AreEqual(2.9d, result);
        }

        [TestMethod]
        public void RoundupShouldReturnAResult()
        {
            var result = _parser.Parse("Roundup(2.99, 1)");
            Assert.AreEqual(3d, result);
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
            var result = _parser.Parse("Count(1,2,2,\"4\")");
            Assert.AreEqual(4d, result);
        }

        [TestMethod]
        public void CountAShouldReturnAResult()
        {
            var result = _parser.Parse("CountA(1,2,2,\"\", \"a\")");
            Assert.AreEqual(4d, result);
        }

        [TestMethod]
        public void CountIfShouldReturnAResult()
        {
            var result = _parser.Parse("CountIf({1;2;2;\"\"}, \"2\")");
            Assert.AreEqual(2d, result);
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
        public void TruncShouldReturnAResult()
        {
            var result = _parser.Parse("Trunc(1.2345)");
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

        [TestMethod]
        public void LnShouldReturnAResult()
        {
            var result = _parser.Parse("Ln(2)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void FactShouldReturnAResult()
        {
            var result = _parser.Parse("Fact(0)");
            Assert.AreEqual(1d, result);
        }

        [TestMethod]
        public void QuotientShouldReturnAResult()
        {
            var result = _parser.Parse("Quotient(5;2)");
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void MedianShouldReturnAResult()
        {
            var result = _parser.Parse("Median(1;2;3)");
            Assert.AreEqual(2d, result);
        }

        [TestMethod]
        public void CountBlankShouldCalculateEmptyCells()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B2"].Value = string.Empty;
                sheet.Cells["A5"].Formula = "COUNTBLANK(A1:B4)";
                sheet.Calculate();
                Assert.AreEqual(7, sheet.Cells["A5"].Value);
            }
        }

        [TestMethod]
        public void DegreesShouldReturnCorrectResult()
        {
            var result = _parser.Parse("DEGREES(0.5)");
            var rounded = Math.Round((double)result, 3);
            Assert.AreEqual(28.648, rounded);
        }

        [TestMethod]
        public void AverateIfsShouldCaluclateResult()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["F4"].Value = 1;
                sheet.Cells["F5"].Value = 2;
                sheet.Cells["F6"].Formula = "2 + 2";
                sheet.Cells["F7"].Value = 4;
                sheet.Cells["F8"].Value = 5;

                sheet.Cells["H4"].Value = 3;
                sheet.Cells["H5"].Value = 3;
                sheet.Cells["H6"].Formula = "2 + 2";
                sheet.Cells["H7"].Value = 4;
                sheet.Cells["H8"].Value = 5;

                sheet.Cells["I4"].Value = 2;
                sheet.Cells["I5"].Value = 3;
                sheet.Cells["I6"].Formula = "2 + 2";
                sheet.Cells["I7"].Value = 5;
                sheet.Cells["I8"].Value = 1;

                sheet.Cells["H9"].Formula = "AVERAGEIFS(F4:F8;H4:H8;\">3\";I4:I8;\"<5\")";
                sheet.Calculate();
                Assert.AreEqual(4.5d, sheet.Cells["H9"].Value);
            }
        }
    }
}
