using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Exceptions;

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
            var args = FunctionsHelper.CreateArgs(ExcelErrorCodes.Value.Code);
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
    }
}
