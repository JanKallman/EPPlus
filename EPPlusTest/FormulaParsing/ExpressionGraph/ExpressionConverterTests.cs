using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Operators;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class ExpressionConverterTests
    {
        private IExpressionConverter _converter;

        [TestInitialize]
        public void Setup()
        {
            _converter = new ExpressionConverter();
        }

        [TestMethod]
        public void ToStringExpressionShouldConvertIntegerExpressionToStringExpression()
        {
            var integerExpression = new IntegerExpression("2");
            var result = _converter.ToStringExpression(integerExpression);
            Assert.IsInstanceOfType(result, typeof(StringExpression));
            Assert.AreEqual("2", result.Compile().Result);
        }

        [TestMethod]
        public void ToStringExpressionShouldCopyOperatorToStringExpression()
        {
            var integerExpression = new IntegerExpression("2");
            integerExpression.Operator = Operator.Plus;
            var result = _converter.ToStringExpression(integerExpression);
            Assert.AreEqual(integerExpression.Operator, result.Operator);
        }

        [TestMethod]
        public void ToStringExpressionShouldConvertDecimalExpressionToStringExpression()
        {
            var decimalExpression = new DecimalExpression("2.5");
            var result = _converter.ToStringExpression(decimalExpression);
            Assert.IsInstanceOfType(result, typeof(StringExpression));
            Assert.AreEqual("2,5", result.Compile().Result);
        }

        [TestMethod]
        public void FromCompileResultShouldCreateIntegerExpressionIfCompileResultIsInteger()
        {
            var compileResult = new CompileResult(1, DataType.Integer);
            var result = _converter.FromCompileResult(compileResult);
            Assert.IsInstanceOfType(result, typeof(IntegerExpression));
            Assert.AreEqual(1d, result.Compile().Result);
        }

        [TestMethod]
        public void FromCompileResultShouldCreateStringExpressionIfCompileResultIsString()
        {
            var compileResult = new CompileResult("abc", DataType.String);
            var result = _converter.FromCompileResult(compileResult);
            Assert.IsInstanceOfType(result, typeof(StringExpression));
            Assert.AreEqual("abc", result.Compile().Result);
        }

        [TestMethod]
        public void FromCompileResultShouldCreateDecimalExpressionIfCompileResultIsDecimal()
        {
            var compileResult = new CompileResult("2.5", DataType.Decimal);
            var result = _converter.FromCompileResult(compileResult);
            Assert.IsInstanceOfType(result, typeof(DecimalExpression));
            Assert.AreEqual(2.5d, result.Compile().Result);
        }

        [TestMethod]
        public void FromCompileResultShouldCreateBooleanExpressionIfCompileResultIsBoolean()
        {
            var compileResult = new CompileResult("true", DataType.Boolean);
            var result = _converter.FromCompileResult(compileResult);
            Assert.IsInstanceOfType(result, typeof(BooleanExpression));
            Assert.IsTrue((bool)result.Compile().Result);
        }
    }
}
