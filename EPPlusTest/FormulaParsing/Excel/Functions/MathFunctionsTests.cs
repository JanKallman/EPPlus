using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Threading;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class MathFunctionsTests
    {
        private ParsingContext _parsingContext;

        [TestInitialize]
        public void Initialize()
        {
            _parsingContext = ParsingContext.Create();
            _parsingContext.Scopes.NewScope(RangeAddress.Empty);
        }

        [TestMethod]
        public void PiShouldReturnPIConstant()
        {
            var expectedValue = (double)Math.Round(Math.PI, 14);
            var func = new Pi();
            var args = FunctionsHelper.CreateArgs(0);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void AbsShouldReturnCorrectResult()
        {
            var expectedValue = 3d;
            var func = new Abs();
            var args = FunctionsHelper.CreateArgs(-3d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceLowerThan0()
        {
            var expectedValue = 22.36d;
            var func = new Ceiling();
            var args = FunctionsHelper.CreateArgs(22.35d, 0.01);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsMinus0point1()
        {
            var expectedValue = -22.4d;
            var func = new Ceiling();
            var args = FunctionsHelper.CreateArgs(-22.35d, -0.1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, System.Math.Round((double)result.Result, 2));
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceIs1()
        {
            var expectedValue = 23d;
            var func = new Ceiling();
            var args = FunctionsHelper.CreateArgs(22.35d, 1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundUpAccordingToParamsSignificanceIs10()
        {
            var expectedValue = 30d;
            var func = new Ceiling();
            var args = FunctionsHelper.CreateArgs(22.35d, 10);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void CeilingShouldRoundTowardsZeroIfSignificanceAndNumberIsNegative()
        {
            var expectedValue = -30d;
            var func = new Ceiling();
            var args = FunctionsHelper.CreateArgs(-22.35d, -10);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void CeilingShouldThrowExceptionIfNumberIsPositiveAndSignificanceIsNegative()
        {
            var expectedValue = 30d;
            var func = new Ceiling();
            var args = FunctionsHelper.CreateArgs(22.35d, -1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedValue, result.Result);
        }

        [TestMethod]
        public void SumShouldCalculate2Plus3AndReturn5()
        {
            var func = new Sum();
            var args = FunctionsHelper.CreateArgs(2, 3);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumShouldCalculateEnumerableOf2Plus5Plus3AndReturn10()
        {
            var func = new Sum();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(2, 5), 3);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(10d, result.Result);
        }

        [TestMethod]
        public void SumShouldIgnoreHiddenValuesWhenIgnoreHiddenValuesIsSet()
        {
            var func = new Sum();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(2, 5), 3, 4);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(10d, result.Result);
        }

        [TestMethod]
        public void SumIfShouldCalculateMatchingValuesOnly()
        {
            var func = new SumIf();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(3, 4, 5), "4");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(4d, result.Result);
        }

        [TestMethod]
        public void SumIfShouldCalculateWithExpression()
        {
            var func = new SumIf();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(3, 4, 5), ">3");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(9d, result.Result);
        }


        [TestMethod, ExpectedException(typeof(ExcelErrorValueException))]
        public void SumIfShouldThrowIfCriteriaIsLargerThan255Chars()
        {
            var longString = "a";
            for (var x = 0; x < 256; x++) { longString = string.Concat(longString, "a"); }
            var func = new SumIf();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(3, 4, 5), longString, (FunctionsHelper.CreateArgs(3, 2, 1)));
            var result = func.Execute(args, _parsingContext);
        }

        [TestMethod]
        public void SumSqShouldCalculateArray()
        {
            var func = new Sumsq();
            var args = FunctionsHelper.CreateArgs(2, 4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(20d, result.Result);
        }

        [TestMethod]
        public void SumSqShouldIncludeTrueAsOne()
        {
            var func = new Sumsq();
            var args = FunctionsHelper.CreateArgs(2, 4, true);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(21d, result.Result);
        }

        [TestMethod]
        public void SumSqShouldNoCountTrueTrueInArray()
        {
            var func = new Sumsq();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(2, 4, true));
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(20d, result.Result);
        }

        [TestMethod]
        public void StdevShouldCalculateCorrectResult()
        {
            var func = new Stdev();
            var args = FunctionsHelper.CreateArgs(1, 3, 5);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void StdevShouldIgnoreHiddenValuesWhenIgnoreHiddenValuesIsSet()
        {
            var func = new Stdev();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(1, 3, 5, 6);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void StdevPShouldCalculateCorrectResult()
        {
            var func = new StdevP();
            var args = FunctionsHelper.CreateArgs(2, 3, 4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(0.8165d, Math.Round((double)result.Result, 5));
        }

        [TestMethod]
        public void StdevPShouldIgnoreHiddenValuesWhenIgnoreHiddenValuesIsSet()
        {
            var func = new StdevP();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(2, 3, 4, 165);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(0.8165d, Math.Round((double)result.Result, 5));
        }

        [TestMethod]
        public void ExpShouldCalculateCorrectResult()
        {
            var func = new Exp();
            var args = FunctionsHelper.CreateArgs(4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(54.59815003d, System.Math.Round((double)result.Result, 8));
        }

        [TestMethod]
        public void MaxShouldCalculateCorrectResult()
        {
            var func = new Max();
            var args = FunctionsHelper.CreateArgs(4, 2, 5, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void MaxShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
        {
            var func = new Max();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(4, 2, 5, 2);
            args.ElementAt(2).SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(4d, result.Result);
        }

        [TestMethod]
        public void MaxaShouldCalculateCorrectResult()
        {
            var func = new Maxa();
            var args = FunctionsHelper.CreateArgs(-1, 0, 1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void MaxaShouldCalculateCorrectResultUsingBool()
        {
            var func = new Maxa();
            var args = FunctionsHelper.CreateArgs(-1, 0, true);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void MaxaShouldCalculateCorrectResultUsingString()
        {
            var func = new Maxa();
            var args = FunctionsHelper.CreateArgs(-1, "test");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(0d, result.Result);
        }

        [TestMethod]
        public void MinShouldCalculateCorrectResult()
        {
            var func = new Min();
            var args = FunctionsHelper.CreateArgs(4, 2, 5, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void MinShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
        {
            var func = new Min();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(4, 2, 5, 3);
            args.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void AverageShouldCalculateCorrectResult()
        {
            var expectedResult = (4d + 2d + 5d + 2d) / 4d;
            var func = new Average();
            var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void AverageShouldCalculateCorrectResultWithEnumerableAndBoolMembers()
        {
            var expectedResult = (4d + 2d + 5d + 2d + 1d) / 5d;
            var func = new Average();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4d, 2d), 5d, 2d, true);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void AverageShouldIgnoreHiddenFieldsIfIgnoreHiddenValuesIsTrue()
        {
            var expectedResult = (4d + 2d + 2d + 1d) / 4d;
            var func = new Average();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4d, 2d), 5d, 2d, true);
            args.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void AverageShouldThrowDivByZeroExcelErrorValueIfEmptyArgs()
        {
            eErrorType errorType = eErrorType.Value;

            var func = new Average();
            var args = new FunctionArgument[0];
            try
            {
                func.Execute(args, _parsingContext);
            }
            catch (ExcelErrorValueException e)
            {
                errorType = e.ErrorValue.Type;
            }
            Assert.AreEqual(eErrorType.Div0, errorType);
        }

        [TestMethod]
        public void AverageAShouldCalculateCorrectResult()
        {
            var expectedResult = (4d + 2d + 5d + 2d) / 4d;
            var func = new AverageA();
            var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void AverageAShouldIncludeTrueAs1()
        {
            var expectedResult = (4d + 2d + 5d + 2d + 1d) / 5d;
            var func = new AverageA();
            var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d, true);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod, ExpectedException(typeof(ExcelErrorValueException))]
        public void AverageAShouldThrowValueExceptionIfNonNumericTextIsSupplied()
        {
            var func = new AverageA();
            var args = FunctionsHelper.CreateArgs(4d, 2d, 5d, 2d, "ABC");
            var result = func.Execute(args, _parsingContext);
        }

        [TestMethod]
        public void AverageAShouldCountValueAs0IfNonNumericTextIsSuppliedInArray()
        {
            var func = new AverageA();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1d, 2d, 3d, "ABC"));
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.5d,result.Result);
        }

        [TestMethod]
        public void AverageAShouldCountNumericStringWithValue()
        {
            var func = new AverageA();
            var args = FunctionsHelper.CreateArgs(4d, 2d, "9");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void RoundShouldReturnCorrectResult()
        {
            var func = new Round();
            var args = FunctionsHelper.CreateArgs(2.3433, 3);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2.343d, result.Result);
        }

        [TestMethod]
        public void RoundShouldReturnCorrectResultWhenNbrOfDecimalsIsNegative()
        {
            var func = new Round();
            var args = FunctionsHelper.CreateArgs(9333, -3);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(9000d, result.Result);
        }

        [TestMethod]
        public void FloorShouldReturnCorrectResultWhenSignificanceIsBetween0And1()
        {
            var func = new Floor();
            var args = FunctionsHelper.CreateArgs(26.75d, 0.1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(26.7d, result.Result);
        }

        [TestMethod]
        public void FloorShouldReturnCorrectResultWhenSignificanceIs1()
        {
            var func = new Floor();
            var args = FunctionsHelper.CreateArgs(26.75d, 1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(26d, result.Result);
        }

        [TestMethod]
        public void FloorShouldReturnCorrectResultWhenSignificanceIsMinus1()
        {
            var func = new Floor();
            var args = FunctionsHelper.CreateArgs(-26.75d, -1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(-26d, result.Result);
        }

        [TestMethod]
        public void RandShouldReturnAValueBetween0and1()
        {
            var func = new Rand();
            var args = new FunctionArgument[0];
            var result1 = func.Execute(args, _parsingContext);
            Assert.IsTrue(((double)result1.Result) > 0 && ((double) result1.Result) < 1);
            var result2 = func.Execute(args, _parsingContext);
            Assert.AreNotEqual(result1.Result, result2.Result, "The two numbers were the same");
            Assert.IsTrue(((double)result2.Result) > 0 && ((double)result2.Result) < 1);
        }

        [TestMethod]
        public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValues()
        {
            var func = new RandBetween();
            var args = FunctionsHelper.CreateArgs(1, 5);
            var result = func.Execute(args, _parsingContext);
            CollectionAssert.Contains(new List<double> { 1d, 2d, 3d, 4d, 5d }, result.Result);
        }

        [TestMethod]
        public void RandBetweenShouldReturnAnIntegerValueBetweenSuppliedValuesWhenLowIsNegative()
        {
            var func = new RandBetween();
            var args = FunctionsHelper.CreateArgs(-5, 0);
            var result = func.Execute(args, _parsingContext);
            CollectionAssert.Contains(new List<double> { 0d, -1d, -2d, -3d, -4d, -5d }, result.Result);
        }

        [TestMethod]
        public void CountShouldReturnNumberOfNumericItems()
        {
            var func = new Count();
            var args = FunctionsHelper.CreateArgs(1d, 2m, 3, new DateTime(2012, 4, 1), "4");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void CountShouldIncludeNumericStringsAndDatesInArray()
        {
            var func = new Count();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1d, 2m, 3, new DateTime(2012, 4, 1), "4"));
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }


        [TestMethod]
        public void CountShouldIncludeEnumerableMembers()
        {
            var func = new Count();
            var args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod, Ignore]
        public void CountShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
        {
            var func = new Count();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
            args.ElementAt(0).SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void CountAShouldReturnNumberOfNonWhitespaceItems()
        {
            var func = new CountA();
            var args = FunctionsHelper.CreateArgs(1d, 2m, 3, new DateTime(2012, 4, 1), "4", null, string.Empty);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void CountAShouldIncludeEnumerableMembers()
        {
            var func = new CountA();
            var args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void CountAShouldIgnoreHiddenValuesIfIgnoreHiddenValuesIsTrue()
        {
            var func = new CountA();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(1d, FunctionsHelper.CreateArgs(12, 13));
            args.ElementAt(0).SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void CountIfShouldReturnNbrOfNumericItemsThatMatch()
        {
            var func = new CountIf();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1d, 2d, 3d), ">1");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void CountIfShouldReturnNbrOfAlphaNumItemsThatMatch()
        {
            var func = new CountIf();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs("Monday", "Tuesday", "Thursday"), "T*day");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod]
        public void ProductShouldMultiplyArguments()
        {
            var func = new Product();
            var args = FunctionsHelper.CreateArgs(2d, 2d, 4d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(16d, result.Result);
        }

        [TestMethod]
        public void ProductShouldHandleEnumerable()
        {
            var func = new Product();
            var args = FunctionsHelper.CreateArgs(2d, 2d, FunctionsHelper.CreateArgs(4d, 2d));
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(32d, result.Result);
        }

        [TestMethod]
        public void ProductShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
        {
            var func = new Product();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(2d, 2d, FunctionsHelper.CreateArgs(4d, 2d));
            args.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(16d, result.Result);
        }

        [TestMethod]
        public void ProductShouldHandleFirstItemIsEnumerable()
        {
            var func = new Product();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4d, 2d), 2d, 2d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(32d, result.Result);
        }

        [TestMethod]
        public void VarShouldReturnCorrectResult()
        {
            var func = new Var();
            var args = FunctionsHelper.CreateArgs(1, 2, 3, 4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.6667d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void VarShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
        {
            var func = new Var();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(1, 2, 3, 4, 9);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.6667d, System.Math.Round((double)result.Result, 4));
        }

        [TestMethod]
        public void VarPShouldReturnCorrectResult()
        {
            var func = new VarP();
            var args = FunctionsHelper.CreateArgs(1, 2, 3, 4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.25d, result.Result);
        }

        [TestMethod]
        public void VarPShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
        {
            var func = new VarP();
            func.IgnoreHiddenValues = true;
            var args = FunctionsHelper.CreateArgs(1, 2, 3, 4, 9);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1.25d, result.Result);
        }

        [TestMethod]
        public void ModShouldReturnCorrectResult()
        {
            var func = new Mod();
            var args = FunctionsHelper.CreateArgs(5, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CosShouldReturnCorrectResult()
        {
            var func = new Cos();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(-0.416146837d, roundedResult);
        }

        [TestMethod]
        public void CosHShouldReturnCorrectResult()
        {
            var func = new Cosh();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(3.762195691, roundedResult);
        }

        [TestMethod]
        public void SinShouldReturnCorrectResult()
        {
            var func = new Sin();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(0.909297427, roundedResult);
        }

        [TestMethod]
        public void SinhShouldReturnCorrectResult()
        {
            var func = new Sinh();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(3.626860408d, roundedResult);
        }

        [TestMethod]
        public void TanShouldReturnCorrectResult()
        {
            var func = new Tan();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(-2.185039863d, roundedResult);
        }

        [TestMethod]
        public void TanhShouldReturnCorrectResult()
        {
            var func = new Tanh();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(0.96402758d, roundedResult);
        }

        [TestMethod]
        public void AtanShouldReturnCorrectResult()
        {
            var func = new Atan();
            var args = FunctionsHelper.CreateArgs(10);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double) result.Result, 9);
            Assert.AreEqual(1.471127674d, roundedResult);
        }

        [TestMethod]
        public void Atan2ShouldReturnCorrectResult()
        {
            var func = new Atan2();
            var args = FunctionsHelper.CreateArgs(1,2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(1.107148718d, roundedResult);
        }

        [TestMethod]
        public void LogShouldReturnCorrectResult()
        {
            var func = new Log();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(0.301029996d, roundedResult);
        }

        [TestMethod]
        public void LogShouldReturnCorrectResultWithBase()
        {
            var func = new Log();
            var args = FunctionsHelper.CreateArgs(2, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void Log10ShouldReturnCorrectResult()
        {
            var func = new Log10();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(0.301029996d, roundedResult);
        }

        [TestMethod]
        public void LnShouldReturnCorrectResult()
        {
            var func = new Ln();
            var args = FunctionsHelper.CreateArgs(5);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 5);
            Assert.AreEqual(1.60944d, roundedResult);
        }

        [TestMethod]
        public void SqrtPiShouldReturnCorrectResult()
        {
            var func = new SqrtPi();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext);
            var roundedResult = Math.Round((double)result.Result, 9);
            Assert.AreEqual(2.506628275d, roundedResult);
        }

        [TestMethod]
        public void SignShouldReturnMinus1IfArgIsNegative()
        {
            var func = new Sign();
            var args = FunctionsHelper.CreateArgs(-2);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(-1d, result);
        }

        [TestMethod]
        public void SignShouldReturn1IfArgIsPositive()
        {
            var func = new Sign();
            var args = FunctionsHelper.CreateArgs(2);
            var result = func.Execute(args, _parsingContext).Result;
            Assert.AreEqual(1d, result);
        }

        [TestMethod]
        public void RounddownShouldReturnCorrectResultWithPositiveNumber()
        {
            var func = new Rounddown();
            var args = FunctionsHelper.CreateArgs(9.999, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(9.99, result.Result);
        }

        [TestMethod]
        public void RounddownShouldHandleNegativeNumber()
        {
            var func = new Rounddown();
            var args = FunctionsHelper.CreateArgs(-9.999, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(-9.99, result.Result);
        }

        [TestMethod]
        public void RounddownShouldHandleNegativeNumDigits()
        {
            var func = new Rounddown();
            var args = FunctionsHelper.CreateArgs(999.999, -2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(900d, result.Result);
        }

        [TestMethod]
        public void RounddownShouldReturn0IfNegativeNumDigitsIsTooLarge()
        {
            var func = new Rounddown();
            var args = FunctionsHelper.CreateArgs(999.999, -4);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(0d, result.Result);
        }

        [TestMethod]
        public void RounddownShouldHandleZeroNumDigits()
        {
            var func = new Rounddown();
            var args = FunctionsHelper.CreateArgs(999.999, 0);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(999d, result.Result);
        }

        [TestMethod]
        public void RoundupShouldReturnCorrectResultWithPositiveNumber()
        {
            var func = new Roundup();
            var args = FunctionsHelper.CreateArgs(9.9911, 3);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(9.992, result.Result);
        }

        [TestMethod]
        public void RoundupShouldHandleNegativeNumDigits()
        {
            var func = new Roundup();
            var args = FunctionsHelper.CreateArgs(99123, -2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(99200d, result.Result);
        }

        [TestMethod]
        public void RoundupShouldHandleZeroNumDigits()
        {
            var func = new Roundup();
            var args = FunctionsHelper.CreateArgs(999.999, 0);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1000d, result.Result);
        }

        [TestMethod]
        public void TruncShouldReturnCorrectResult()
        {
            var func = new Trunc();
            var args = FunctionsHelper.CreateArgs(99.99);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(99d, result.Result);
        }

        [TestMethod]
        public void FactShouldRoundDownAndReturnCorrectResult()
        {
            var func = new Fact();
            var args = FunctionsHelper.CreateArgs(5.99);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(120d, result.Result);
        }

        [TestMethod, ExpectedException(typeof (ExcelErrorValueException))]
        public void FactShouldThrowWhenNegativeNumber()
        {
            var func = new Fact();
            var args = FunctionsHelper.CreateArgs(-1);
            func.Execute(args, _parsingContext);
        }

        [TestMethod]
        public void QuotientShouldReturnCorrectResult()
        {
            var func = new Quotient();
            var args = FunctionsHelper.CreateArgs(5, 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod, ExpectedException(typeof(ExcelErrorValueException))]
        public void QuotientShouldThrowWhenDenomIs0()
        {
            var func = new Quotient();
            var args = FunctionsHelper.CreateArgs(1, 0);
            func.Execute(args, _parsingContext);
        }

        [TestMethod]
        public void LargeShouldReturnTheLargestNumberIf1()
        {
            var func = new Large();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1, 2, 3), 1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void LargeShouldReturnTheSecondLargestNumberIf2()
        {
            var func = new Large();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod, ExpectedException(typeof(ExcelErrorValueException))]
        public void LargeShouldThrowIfIndexOutOfBounds()
        {
            var func = new Large();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 6);
            var result = func.Execute(args, _parsingContext);
        }

        [TestMethod]
        public void SmallShouldReturnTheSmallestNumberIf1()
        {
            var func = new Small();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1, 2, 3), 1);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void SmallShouldReturnTheSecondSmallestNumberIf2()
        {
            var func = new Small();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2d, result.Result);
        }

        [TestMethod, ExpectedException(typeof(ExcelErrorValueException))]
        public void SmallShouldThrowIfIndexOutOfBounds()
        {
            var func = new Small();
            var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4, 1, 2, 3), 6);
            var result = func.Execute(args, _parsingContext);
        }
    }
}
