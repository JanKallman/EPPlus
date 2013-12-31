using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class InformationFunctionsTests
    {
        private ParsingContext _context;

        [TestInitialize]
        public void Setup()
        {
            _context = ParsingContext.Create();
        }

        [TestMethod]
        public void IsBlankShouldReturnTrueIfFirstArgIsNull()
        {
            var func = new IsBlank();
            var args = FunctionsHelper.CreateArgs(new object[]{null});
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsBlankShouldReturnTrueIfFirstArgIsEmptyString()
        {
            var func = new IsBlank();
            var args = FunctionsHelper.CreateArgs(string.Empty);
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsNumberShouldReturnTrueWhenArgIsNumeric()
        {
            var func = new IsNumber();
            var args = FunctionsHelper.CreateArgs(1d);
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsNumberShouldReturnfalseWhenArgIsNonNumeric()
        {
            var func = new IsNumber();
            var args = FunctionsHelper.CreateArgs("1");
            var result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void IsErrorShouldReturnTrueIfArgIsAnErrorCode()
        {
            var args = FunctionsHelper.CreateArgs(ExcelErrorValue.Parse("#DIV/0!"));
            var func = new IsError();
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsErrorShouldReturnFalseIfArgIsNotAnError()
        {
            var args = FunctionsHelper.CreateArgs("A", 1);
            var func = new IsError();
            var result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void IsTextShouldReturnTrueWhenFirstArgIsAString()
        {
            var args = FunctionsHelper.CreateArgs("1");
            var func = new IsText();
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsTextShouldReturnFalseWhenFirstArgIsNotAString()
        {
            var args = FunctionsHelper.CreateArgs(1);
            var func = new IsText();
            var result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);
        }

        [TestMethod]
        public void IsOddShouldReturnCorrectResult()
        {
            var args = FunctionsHelper.CreateArgs(3.123);
            var func = new IsOdd();
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsEvenShouldReturnCorrectResult()
        {
            var args = FunctionsHelper.CreateArgs(4.123);
            var func = new IsEven();
            var result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }

        [TestMethod]
        public void IsLogicalShouldReturnCorrectResult()
        {
            var func = new IsLogical();

            var args = FunctionsHelper.CreateArgs(1);
            var result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);

            args = FunctionsHelper.CreateArgs("true");
            result = func.Execute(args, _context);
            Assert.IsFalse((bool)result.Result);

            args = FunctionsHelper.CreateArgs(false);
            result = func.Execute(args, _context);
            Assert.IsTrue((bool)result.Result);
        }
    }
}
