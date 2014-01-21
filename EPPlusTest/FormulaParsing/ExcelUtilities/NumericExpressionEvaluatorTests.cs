using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest
{
    [TestClass]
    public class NumericExpressionEvaluatorTests
    {
        private NumericExpressionEvaluator _evaluator;

        [TestInitialize]
        public void Setup()
        {
            _evaluator = new NumericExpressionEvaluator();
        }

        [TestMethod]
        public void EvaluateShouldReturnTrueIfOperandsAreEqual()
        {
            var result = _evaluator.Evaluate("1", "1");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateShouldReturnTrueIfOperandsAreMatchingButDifferentTypes()
        {
            var result = _evaluator.Evaluate(1d, "1");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateShouldEvaluateOperator()
        {
            var result = _evaluator.Evaluate(1d, "<2");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateShouldEvaluateNumericString()
        {
            var result = _evaluator.Evaluate("1", ">0");
            Assert.IsTrue(result);
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void EvaluateShouldThrowIfOperatorIsNotBoolean()
        {
            var result = _evaluator.Evaluate(1d, "+1");
        }
    }
}
