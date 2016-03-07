using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions.ExcelRanges
{
    [TestClass]
    public class TextExcelRangeTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;
        private CultureInfo _currentCulture;

        [TestInitialize]
        public void Initialize()
        {
            _currentCulture = CultureInfo.CurrentCulture;
            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add("Test");

            _worksheet.Cells["A1"].Value = 1;
            _worksheet.Cells["A2"].Value = 3;
            _worksheet.Cells["A3"].Value = 6;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
            Thread.CurrentThread.CurrentCulture = _currentCulture;
        }

        [TestMethod]
        public void ExactShouldReturnTrueWhenEqualValues()
        {
            _worksheet.Cells["A2"].Value = 1d;
            _worksheet.Cells["A4"].Formula = "EXACT(A1,A2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.IsTrue((bool)result);
        }

        [TestMethod]
        public void FindShouldReturnIndex()
        {
            _worksheet.Cells["A1"].Value = "hopp";
            _worksheet.Cells["A2"].Value = "hej hopp";
            _worksheet.Cells["A4"].Formula = "Find(A1,A2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(5, result);
        }

        [TestMethod]
        public void ValueShouldHandleStringWithIntegers()
        {
            _worksheet.Cells["A1"].Value = "12";
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(12d, result);
        }

        [TestMethod]
        public void ValueShouldHandle1000delimiter()
        {
            var delimiter = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
            var val = $"5{delimiter}000";
            _worksheet.Cells["A1"].Value = val;
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(5000d, result);
        }

        [TestMethod]
        public void ValueShouldHandle1000DelimiterAndDecimal()
        {
            var delimiter = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
            var decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            var val = $"5{delimiter}000{decimalSeparator}123";
            _worksheet.Cells["A1"].Value = val;
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(5000.123d, result);
        }

        [TestMethod]
        public void ValueShouldHandlePercent()
        {
            var val = $"20%";
            _worksheet.Cells["A1"].Value = val;
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(0.2d, result);
        }

        [TestMethod]
        public void ValueShouldHandleScientificNotation()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            _worksheet.Cells["A1"].Value = "1.2345E-02";
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(0.012345d, result);
        }

        [TestMethod]
        public void ValueShouldHandleDate()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var date = new DateTime(2015, 12, 31);
            _worksheet.Cells["A1"].Value = date.ToString(CultureInfo.CurrentCulture);
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(date.ToOADate(), result);
        }

        [TestMethod]
        public void ValueShouldHandleTime()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var date = new DateTime(2015, 12, 31);
            var date2 = new DateTime(2015, 12, 31, 12, 00, 00);
            var ts = date2.Subtract(date);
            _worksheet.Cells["A1"].Value = ts.ToString();
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.AreEqual(0.5, result);
        }
    }
}
