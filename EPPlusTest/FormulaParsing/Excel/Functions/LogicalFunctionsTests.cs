using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class LogicalFunctionsTests
    {
        private ParsingContext _parsingContext = ParsingContext.Create();

        [TestMethod]
        public void IfShouldReturnCorrectResult()
        {
            var func = new If();
            var args = FunctionsHelper.CreateArgs(true, "A", "B");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual("A", result.Result);
        }

        [TestMethod]
        public void NotShouldReturnFalseIfArgumentIsTrue()
        {
            var func = new Not();
            var args = FunctionsHelper.CreateArgs(true);
            var result = func.Execute(args, _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void NotShouldReturnTrueIfArgumentIs0()
        {
            var func = new Not();
            var args = FunctionsHelper.CreateArgs(0);
            var result = func.Execute(args, _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void NotShouldReturnFalseIfArgumentIs1()
        {
            var func = new Not();
            var args = FunctionsHelper.CreateArgs(1);
            var result = func.Execute(args, _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void NotShouldHandleExcelReference()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = false;
                sheet.Cells["A2"].Formula = "NOT(A1)";
                sheet.Calculate();
                Assert.IsTrue((bool)sheet.Cells["A2"].Value);
            }
        }

        [TestMethod]
        public void AndShouldReturnTrueIfAllArgumentsAreTrue()
        {
            var func = new And();
            var args = FunctionsHelper.CreateArgs(true, true, true);
            var result = func.Execute(args, _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void AndShouldReturnTrueIfAllArgumentsAreTrueOr1()
        {
            var func = new And();
            var args = FunctionsHelper.CreateArgs(true, true, 1, true, 1);
            var result = func.Execute(args, _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void AndShouldReturnFalseIfOneArgumentIsFalse()
        {
            var func = new And();
            var args = FunctionsHelper.CreateArgs(true, false, true);
            var result = func.Execute(args, _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void AndShouldReturnFalseIfOneArgumentIs0()
        {
            var func = new And();
            var args = FunctionsHelper.CreateArgs(true, 0, true);
            var result = func.Execute(args, _parsingContext);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void OrShouldReturnTrueIfOneArgumentIsTrue()
        {
            var func = new Or();
            var args = FunctionsHelper.CreateArgs(true, false, false);
            var result = func.Execute(args, _parsingContext);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IfErrorShouldReturnSecondArgIfCriteriaEvaluatesAsAnError()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFERROR(0/0, \"hello\")";
                s1.Calculate();
                Assert.AreEqual("hello", s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfErrorShouldReturnSecondArgIfCriteriaEvaluatesAsAnError2()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFERROR(A2, \"hello\")";
                s1.Cells["A2"].Formula = "23/0";
                s1.Calculate();
                Assert.AreEqual("hello", s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfErrorShouldReturnResultOfFormulaIfNoError()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFERROR(A2, \"hello\")";
                s1.Cells["A2"].Value = "hi there";
                s1.Calculate();
                Assert.AreEqual("hi there", s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfNaShouldReturnSecondArgIfCriteriaEvaluatesAsAnError2()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFERROR(A2, \"hello\")";
                s1.Cells["A2"].Value = ExcelErrorValue.Create(eErrorType.NA);
                s1.Calculate();
                Assert.AreEqual("hello", s1.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfNaShouldReturnResultOfFormulaIfNoError()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFNA(A2, \"hello\")";
                s1.Cells["A2"].Value = "hi there";
                s1.Calculate();
                Assert.AreEqual("hi there", s1.Cells["A1"].Value);
            }
        }
    }
}
