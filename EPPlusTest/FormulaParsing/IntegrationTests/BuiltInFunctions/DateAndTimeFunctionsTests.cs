﻿using System;
using System.Text;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using FakeItEasy;
using System.IO;
using System.Threading;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class DateAndTimeFunctionsTests : FormulaParserTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            var excelDataProvider = A.Fake<ExcelDataProvider>();
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
            var result = _parser.Parse("Second(\"10:12:14\")");
            Assert.AreEqual(14, result);
        }

        [TestMethod]
        public void MinuteShouldReturnCorrectResultWhenParsingString()
        {
            var result = _parser.Parse("Minute(\"10:12:14 AM\")");
            Assert.AreEqual(12, result);
        }

        [TestMethod]
        public void HourShouldReturnCorrectResultWhenParsingString()
        {
            var result = _parser.Parse("Hour(\"10:12:14\")");
            Assert.AreEqual(10, result);
        }

        [TestMethod]
        public void Day360ShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Days360(Date(2012, 4, 2), Date(2012, 5, 2))");
            Assert.AreEqual(30, result);
        }

        [TestMethod]
        public void YearfracShouldReturnAResult()
        {
            var result = _parser.Parse("Yearfrac(Date(2012, 4, 2), Date(2012, 5, 2))");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void IsoWeekNumShouldReturnAResult()
        {
            var result = _parser.Parse("IsoWeekNum(Date(2012, 4, 2))");
            Assert.IsInstanceOfType(result, typeof(int));
        }

        [TestMethod]
        public void EomonthShouldReturnAResult()
        {
            var result = _parser.Parse("Eomonth(Date(2013, 2, 2), 3)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void WorkdayShouldReturnAResult()
        {
            var result = _parser.Parse("Workday(Date(2013, 2, 2), 3)");
            Assert.IsInstanceOfType(result, typeof(double));
        }

        [TestMethod]
        public void DateNotEqualToStringShouldBeTrue()
        {
            var result = _parser.Parse("TODAY() <> \"\"");
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void Calculation5()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = "John";
            ws.Cells["B1"].Value = "Doe";
            ws.Cells["C1"].Formula = "B1&\", \"&A1";
            ws.Calculate();
            Assert.AreEqual("Doe, John", ws.Cells["C1"].Value);
        }

        [TestMethod]
        public void HourWithExcelReference()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = new DateTime(2014, 1, 1, 10, 11, 12).ToOADate();
            ws.Cells["B1"].Formula = "HOUR(A1)";
            ws.Calculate();
            Assert.AreEqual(10, ws.Cells["B1"].Value);
        }

        [TestMethod]
        public void MinuteWithExcelReference()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = new DateTime(2014, 1, 1, 10, 11, 12).ToOADate();
            ws.Cells["B1"].Formula = "MINUTE(A1)";
            ws.Calculate();
            Assert.AreEqual(11, ws.Cells["B1"].Value);
        }

        [TestMethod]
        public void SecondWithExcelReference()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = new DateTime(2014, 1, 1, 10, 11, 12).ToOADate();
            ws.Cells["B1"].Formula = "SECOND(A1)";
            ws.Calculate();
            Assert.AreEqual(12, ws.Cells["B1"].Value);
        }
#if (!Core)
        [TestMethod]
        public void DateValueTest1()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = "21 JAN 2015";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.AreEqual(new DateTime(2015, 1, 21).ToOADate(), ws.Cells["B1"].Value);
        }

        [TestMethod]
        public void DateValueTestWithoutYear()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var currentYear = DateTime.Now.Year;
            ws.Cells["A1"].Value = "21 JAN";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.AreEqual(new DateTime(currentYear, 1, 21).ToOADate(), ws.Cells["B1"].Value);
        }

        [TestMethod]
        public void DateValueTestWithTwoDigitYear()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var expectedYear = 1930;
            ws.Cells["A1"].Value = "01/01/30";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.AreEqual(new DateTime(expectedYear, 1, 1).ToOADate(), ws.Cells["B1"].Value);
        }

        [TestMethod]
        public void DateValueTestWithTwoDigitYear2()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var expectedYear = 2029;
            ws.Cells["A1"].Value = "01/01/29";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.AreEqual(new DateTime(expectedYear, 1, 1).ToOADate(), ws.Cells["B1"].Value);
        }


        [TestMethod]
        public void TimeValueTestPm()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var currentYear = DateTime.Now.Year;
            ws.Cells["A1"].Value = "2:23 pm";
            ws.Cells["B1"].Formula = "TimeValue(A1)";
            ws.Calculate();
            var result = (double) ws.Cells["B1"].Value;
            Assert.AreEqual(0.599, Math.Round(result, 3));
        }


        [TestMethod]
        public void TimeValueTestFullDate()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var currentYear = DateTime.Now.Year;
            ws.Cells["A1"].Value = "01/01/2011 02:23";
            ws.Cells["B1"].Formula = "TimeValue(A1)";
            ws.Calculate();
            var result = (double)ws.Cells["B1"].Value;
            Assert.AreEqual(0.099, Math.Round(result, 3));
        }
#endif
    }
}
