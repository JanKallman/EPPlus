using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Threading;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class DateTimeFunctionsTests
    {
        private ParsingContext _parsingContext = ParsingContext.Create();

        private double GetTime(int hour, int minute, int second)
        {
            var secInADay = DateTime.Today.AddDays(1).Subtract(DateTime.Today).TotalSeconds;
            var secondsOfExample = (double)(hour * 60 * 60 + minute * 60 + second);
            return secondsOfExample / secInADay;
        }
        [TestMethod]
        public void DateFunctionShouldReturnADate()
        {
            var func = new Date();
            var args = FunctionsHelper.CreateArgs(2012, 4, 3);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(DataType.Date, result.DataType);
        }

        [TestMethod]
        public void DateFunctionShouldReturnACorrectDate()
        {
            var expectedDate = new DateTime(2012, 4, 3);
            var func = new Date();
            var args = FunctionsHelper.CreateArgs(2012, 4, 3);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedDate.ToOADate(), result.Result);
        }

        [TestMethod]
        public void DateFunctionShouldMonthFromPrevYearIfMonthIsNegative()
        {
            var expectedDate = new DateTime(2011, 11, 3);
            var func = new Date();
            var args = FunctionsHelper.CreateArgs(2012, -1, 3);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(expectedDate.ToOADate(), result.Result);
        }

        [TestMethod]
        public void NowFunctionShouldReturnNow()
        {
            var startTime = DateTime.Now;
            Thread.Sleep(1);
            var func = new Now();
            var args = new FunctionArgument[0];
            var result = func.Execute(args, _parsingContext);
            Thread.Sleep(1);
            var endTime = DateTime.Now;
            var resultDate = DateTime.FromOADate((double)result.Result);
            Assert.IsTrue(resultDate > startTime && resultDate < endTime);
        }

        [TestMethod]
        public void TodayFunctionShouldReturnTodaysDate()
        {
            var func = new Today();
            var args = new FunctionArgument[0];
            var result = func.Execute(args, _parsingContext);
            var resultDate = DateTime.FromOADate((double)result.Result);
            Assert.AreEqual(DateTime.Now.Date, resultDate);
        }

        [TestMethod]
        public void DayShouldReturnDayInMonth()
        {
            var date = new DateTime(2012, 3, 12);
            var func = new Day();
            var args = FunctionsHelper.CreateArgs(date.ToOADate());
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(12, result.Result);
        }

        [TestMethod]
        public void DayShouldReturnMonthOfYearWithStringParam()
        {
            var date = new DateTime(2012, 3, 12);
            var func = new Day();
            var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-12"), _parsingContext);
            Assert.AreEqual(12, result.Result);
        }

        [TestMethod]
        public void MonthShouldReturnMonthOfYear()
        {
            var date = new DateTime(2012, 3, 12);
            var func = new Month();
            var result = func.Execute(FunctionsHelper.CreateArgs(date.ToOADate()), _parsingContext);
            Assert.AreEqual(3, result.Result);
        }

        [TestMethod]
        public void MonthShouldReturnMonthOfYearWithStringParam()
        {
            var date = new DateTime(2012, 3, 12);
            var func = new Month();
            var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-12"), _parsingContext);
            Assert.AreEqual(3, result.Result);
        }

        [TestMethod]
        public void YearShouldReturnCorrectYear()
        {
            var date = new DateTime(2012, 3, 12);
            var func = new Year();
            var result = func.Execute(FunctionsHelper.CreateArgs(date.ToOADate()), _parsingContext);
            Assert.AreEqual(2012, result.Result);
        }

        [TestMethod]
        public void YearShouldReturnCorrectYearWithStringParam()
        {
            var date = new DateTime(2012, 3, 12);
            var func = new Year();
            var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-12"), _parsingContext);
            Assert.AreEqual(2012, result.Result);
        }

        [TestMethod]
        public void TimeShouldReturnACorrectSerialNumber()
        {
            var expectedResult = GetTime(10, 11, 12);
            var func = new Time();
            var result = func.Execute(FunctionsHelper.CreateArgs(10, 11, 12), _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);  
        }

        [TestMethod]
        public void TimeShouldParseStringCorrectly()
        {
            var expectedResult = GetTime(10, 11, 12);
            var func = new Time();
            var result = func.Execute(FunctionsHelper.CreateArgs("10:11:12"), _parsingContext);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void TimeShouldThrowExceptionIfSecondsIsOutOfRange()
        {
            var func = new Time();
            var result = func.Execute(FunctionsHelper.CreateArgs(10, 11, 60), _parsingContext);
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void TimeShouldThrowExceptionIfMinuteIsOutOfRange()
        {
            var func = new Time();
            var result = func.Execute(FunctionsHelper.CreateArgs(10, 60, 12), _parsingContext);
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void TimeShouldThrowExceptionIfHourIsOutOfRange()
        {
            var func = new Time();
            var result = func.Execute(FunctionsHelper.CreateArgs(24, 12, 12), _parsingContext);
        }

        [TestMethod]
        public void HourShouldReturnCorrectResult()
        {
            var func = new Hour();
            var result = func.Execute(FunctionsHelper.CreateArgs(GetTime(9, 13, 14)), _parsingContext);
            Assert.AreEqual(9, result.Result);

            result = func.Execute(FunctionsHelper.CreateArgs(GetTime(23, 13, 14)), _parsingContext);
            Assert.AreEqual(23, result.Result);
        }

        [TestMethod]
        public void MinuteShouldReturnCorrectResult()
        {
            var func = new Minute();
            var result = func.Execute(FunctionsHelper.CreateArgs(GetTime(9, 14, 14)), _parsingContext);
            Assert.AreEqual(14, result.Result);

            result = func.Execute(FunctionsHelper.CreateArgs(GetTime(9, 55, 14)), _parsingContext);
            Assert.AreEqual(55, result.Result);
        }

        [TestMethod]
        public void SecondShouldReturnCorrectResult()
        {
            var func = new Second();
            var result = func.Execute(FunctionsHelper.CreateArgs(GetTime(9, 14, 17)), _parsingContext);
            Assert.AreEqual(17, result.Result);
        }

        [TestMethod]
        public void SecondShouldReturnCorrectResultWithStringArgument()
        {
            var func = new Second();
            var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-27 10:11:12"), _parsingContext);
            Assert.AreEqual(12, result.Result);
        }

        [TestMethod]
        public void MinuteShouldReturnCorrectResultWithStringArgument()
        {
            var func = new Minute();
            var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-27 10:11:12"), _parsingContext);
            Assert.AreEqual(11, result.Result);
        }

        [TestMethod]
        public void HourShouldReturnCorrectResultWithStringArgument()
        {
            var func = new Hour();
            var result = func.Execute(FunctionsHelper.CreateArgs("2012-03-27 10:11:12"), _parsingContext);
            Assert.AreEqual(10, result.Result);
        }

        [TestMethod]
        public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs1()
        {
            var func = new Weekday();
            var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 1), _parsingContext);
            Assert.AreEqual(1, result.Result);
        }

        [TestMethod]
        public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs2()
        {
            var func = new Weekday();
            var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 2), _parsingContext);
            Assert.AreEqual(7, result.Result);
        }

        [TestMethod]
        public void WeekdayShouldReturnCorrectResultForASundayWhenReturnTypeIs3()
        {
            var func = new Weekday();
            var result = func.Execute(FunctionsHelper.CreateArgs(new DateTime(2012, 4, 1).ToOADate(), 3), _parsingContext);
            Assert.AreEqual(6, result.Result);
        }

        [TestMethod]
        public void WeekNumShouldReturnCorrectResult()
        {
            var func = new Weeknum();
            var dt1 = new DateTime(2012, 12, 31).ToOADate();
            var dt2 = new DateTime(2012, 1, 1).ToOADate();
            var dt3 = new DateTime(2013, 1, 20).ToOADate();

            var r1 = func.Execute(FunctionsHelper.CreateArgs(dt1), _parsingContext);
            var r2 = func.Execute(FunctionsHelper.CreateArgs(dt2), _parsingContext);
            var r3 = func.Execute(FunctionsHelper.CreateArgs(dt3, 2), _parsingContext);

            Assert.AreEqual(53, r1.Result, "r1.Result was not 53, but " + r1.Result.ToString());
            Assert.AreEqual(1, r2.Result, "r2.Result was not 1, but " + r2.Result.ToString());
            Assert.AreEqual(3, r3.Result, "r3.Result was not 3, but " + r3.Result.ToString());
        }

        [TestMethod]
        public void EdateShouldReturnCorrectResult()
        {
            var func = new Edate();

            var dt1arg = new DateTime(2012, 1, 31).ToOADate();
            var dt2arg = new DateTime(2013, 1, 1).ToOADate();
            var dt3arg = new DateTime(2013, 2, 28).ToOADate();

            var r1 = func.Execute(FunctionsHelper.CreateArgs(dt1arg, 1), _parsingContext);
            var r2 = func.Execute(FunctionsHelper.CreateArgs(dt2arg, -1), _parsingContext);
            var r3 = func.Execute(FunctionsHelper.CreateArgs(dt3arg, 2), _parsingContext);

            var dt1 = DateTime.FromOADate((double) r1.Result);
            var dt2 = DateTime.FromOADate((double)r2.Result);
            var dt3 = DateTime.FromOADate((double)r3.Result);

            var exp1 = new DateTime(2012, 2, 29);
            var exp2 = new DateTime(2012, 12, 1);
            var exp3 = new DateTime(2013, 4, 28);

            Assert.AreEqual(exp1, dt1, "dt1 was not " + exp1.ToString("yyyy-MM-dd") + ", but " + dt1.ToString("yyyy-MM-dd"));
            Assert.AreEqual(exp2, dt2, "dt1 was not " + exp2.ToString("yyyy-MM-dd") + ", but " + dt2.ToString("yyyy-MM-dd"));
            Assert.AreEqual(exp3, dt3, "dt1 was not " + exp3.ToString("yyyy-MM-dd") + ", but " + dt3.ToString("yyyy-MM-dd"));
        }

        [TestMethod]
        public void Days360ShouldReturnCorrectResultWithNoMethodSpecified2()
        {
            var func = new Days360();

            var dt1arg = new DateTime(2013, 1, 1).ToOADate();
            var dt2arg = new DateTime(2013, 3, 31).ToOADate();

            var result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg), _parsingContext);

            Assert.AreEqual(90, result.Result);
        }

        [TestMethod]
        public void Days360ShouldReturnCorrectResultWithEuroMethodSpecified()
        {
            var func = new Days360();

            var dt1arg = new DateTime(2013, 1, 1).ToOADate();
            var dt2arg = new DateTime(2013, 3, 31).ToOADate();

            var result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, true), _parsingContext);

            Assert.AreEqual(89, result.Result);
        }

        [TestMethod]
        public void Days360ShouldHandleFebWithEuroMethodSpecified()
        {
            var func = new Days360();

            var dt1arg = new DateTime(2012, 2, 28).ToOADate();
            var dt2arg = new DateTime(2013, 2, 28).ToOADate();

            var result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, true), _parsingContext);

            Assert.AreEqual(360, result.Result);
        }

        [TestMethod]
        public void Days360ShouldHandleFebWithUsMethodSpecified()
        {
            var func = new Days360();

            var dt1arg = new DateTime(2012, 2, 28).ToOADate();
            var dt2arg = new DateTime(2013, 2, 28).ToOADate();

            var result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, false), _parsingContext);

            Assert.AreEqual(358, result.Result);
        }

        [TestMethod]
        public void Days360ShouldHandleFebWithUsMethodSpecified2()
        {
            var func = new Days360();

            var dt1arg = new DateTime(2013, 2, 28).ToOADate();
            var dt2arg = new DateTime(2013, 3, 31).ToOADate();

            var result = func.Execute(FunctionsHelper.CreateArgs(dt1arg, dt2arg, false), _parsingContext);

            Assert.AreEqual(30, result.Result);
        }
    }
}
