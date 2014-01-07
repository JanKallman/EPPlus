using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Calculation;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using Rhino.Mocks;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class DateAndTimeFunctionsTests : FormulaParserTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            var excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(excelDataProvider);
        }

        [TestMethod]
        public void DateShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Date(2012, 2, 2)");
            Assert.AreEqual(new DateTime(2012, 2, 2).ToOADate(), result);
        }

        [TestMethod]
        public void DateShouldHandleCellReference()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2012d;
                sheet.Cells["A2"].Formula = "Date(A1, 2, 2)";
                sheet.Calculate();
                var result = sheet.Cells["A2"].Value;
                Assert.AreEqual(new DateTime(2012, 2, 2).ToOADate(), result);
            }

        }

        [TestMethod]
        public void TodayShouldReturnAResult()
        {
            var result = _parser.Parse("Today()");
            Assert.IsInstanceOfType(DateTime.FromOADate((double)result), typeof(DateTime));
        }

        [TestMethod]
        public void NowShouldReturnAResult()
        {
            var result = _parser.Parse("now()");
            Assert.IsInstanceOfType(DateTime.FromOADate((double)result), typeof(DateTime));
        }

        [TestMethod]
        public void DayShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Day(Date(2012, 4, 2))");
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void MonthShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Month(Date(2012, 4, 2))");
            Assert.AreEqual(4, result);
        }

        [TestMethod]
        public void YearShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Year(Date(2012, 2, 2))");
            Assert.AreEqual(2012, result);
        }

        [TestMethod]
        public void TimeShouldReturnCorrectResult()
        {
            var expectedResult = ((double)(12 * 60 * 60 + 13 * 60 + 14))/((double)(24 * 60 * 60));
            var result = _parser.Parse("Time(12, 13, 14)");
            Assert.AreEqual(expectedResult, result);
        }

        [TestMethod]
        public void HourShouldReturnCorrectResult()
        {
            var result = _parser.Parse("HOUR(Time(12, 13, 14))");
            Assert.AreEqual(12, result);
        }

        [TestMethod]
        public void MinuteShouldReturnCorrectResult()
        {
            var result = _parser.Parse("minute(Time(12, 13, 14))");
            Assert.AreEqual(13, result);
        }

        [TestMethod]
        public void SecondShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Second(Time(12, 13, 59))");
            Assert.AreEqual(59, result);
        }

        [TestMethod]
        public void SecondShouldReturnCorrectResultWhenParsingString()
        {
            var result = _parser.Parse("Second('10:12:14')");
            Assert.AreEqual(14, result);
        }

        [TestMethod]
        public void MinuteShouldReturnCorrectResultWhenParsingString()
        {
            var result = _parser.Parse("Minute('10:12:14 AM')");
            Assert.AreEqual(12, result);
        }

        [TestMethod]
        public void HourShouldReturnCorrectResultWhenParsingString()
        {
            var result = _parser.Parse("Hour('10:12:14')");
            Assert.AreEqual(10, result);
        }

        [TestMethod]
        public void Day360ShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Days360(Date(2012, 4, 2), Date(2012, 5, 2))");
            Assert.AreEqual(30, result);
        }
    }
}
