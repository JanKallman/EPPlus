using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class SubtotalTests
    {
        private ParsingContext _context;

        [TestInitialize]
        public void Setup()
        {
            _context = ParsingContext.Create();
            _context.Scopes.NewScope(RangeAddress.Empty);
        }

        [TestMethod, ExpectedException(typeof(ExcelErrorValueException))]
        public void ShouldThrowIfInvalidFuncNumber()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(139, 1);
            func.Execute(args, _context);
        }

        [TestMethod]
        public void ShouldCalculateAverageWhenCalcTypeIs1()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(1, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(30d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateCountWhenCalcTypeIs2()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(2, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateCountAWhenCalcTypeIs3()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(3, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateMaxWhenCalcTypeIs4()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(4, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(50d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateMinWhenCalcTypeIs5()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(5, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(10d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateProductWhenCalcTypeIs6()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(6, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(12000000d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateStdevWhenCalcTypeIs7()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(7, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            var resultRounded = Math.Round((double)result.Result, 5);
            Assert.AreEqual(15.81139d, resultRounded);
        }

        [TestMethod]
        public void ShouldCalculateStdevPWhenCalcTypeIs8()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(8, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            var resultRounded = Math.Round((double)result.Result, 8);
            Assert.AreEqual(14.14213562, resultRounded);
        }

        [TestMethod]
        public void ShouldCalculateSumWhenCalcTypeIs9()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(9, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(150d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateVarWhenCalcTypeIs10()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(10, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(250d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateVarPWhenCalcTypeIs11()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(11, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(200d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateAverageWhenCalcTypeIs101()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(101, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(30d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateCountWhenCalcTypeIs102()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(102, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateCountAWhenCalcTypeIs103()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(103, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateMaxWhenCalcTypeIs104()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(104, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(50d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateMinWhenCalcTypeIs105()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(105, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(10d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateProductWhenCalcTypeIs106()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(106, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(12000000d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateStdevWhenCalcTypeIs107()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(107, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            var resultRounded = Math.Round((double)result.Result, 5);
            Assert.AreEqual(15.81139d, resultRounded);
        }

        [TestMethod]
        public void ShouldCalculateStdevPWhenCalcTypeIs108()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(108, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            var resultRounded = Math.Round((double)result.Result, 8);
            Assert.AreEqual(14.14213562, resultRounded);
        }

        [TestMethod]
        public void ShouldCalculateSumWhenCalcTypeIs109()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(109, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(150d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateVarWhenCalcTypeIs110()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(110, 10, 20, 30, 40, 50, 51);
            args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
            var result = func.Execute(args, _context);
            Assert.AreEqual(250d, result.Result);
        }

        [TestMethod]
        public void ShouldCalculateVarPWhenCalcTypeIs111()
        {
            var func = new Subtotal();
            var args = FunctionsHelper.CreateArgs(111, 10, 20, 30, 40, 50);
            var result = func.Execute(args, _context);
            Assert.AreEqual(200d, result.Result);
        }
    }
}
