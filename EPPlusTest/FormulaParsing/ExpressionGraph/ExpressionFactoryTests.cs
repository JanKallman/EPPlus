using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class ExpressionFactoryTests
    {
        private IExpressionFactory _factory;
        private ParsingContext _parsingContext;

        [TestInitialize]
        public void Setup()
        {
            _parsingContext = ParsingContext.Create();
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            _factory = new ExpressionFactory(provider, _parsingContext);
        }

        [TestMethod]
        public void ShouldReturnIntegerExpressionWhenTokenIsInteger()
        {
            var token = new Token("2", TokenType.Integer);
            var expression = _factory.Create(token);
            Assert.IsInstanceOfType(expression, typeof(IntegerExpression));
        }

        [TestMethod]
        public void ShouldReturnBooleanExpressionWhenTokenIsBoolean()
        {
            var token = new Token("true", TokenType.Boolean);
            var expression = _factory.Create(token);
            Assert.IsInstanceOfType(expression, typeof(BooleanExpression));
        }

        [TestMethod]
        public void ShouldReturnDecimalExpressionWhenTokenIsDecimal()
        {
            var token = new Token("2.5", TokenType.Decimal);
            var expression = _factory.Create(token);
            Assert.IsInstanceOfType(expression, typeof(DecimalExpression));
        }

        [TestMethod]
        public void ShouldReturnExcelRangeExpressionWhenTokenIsExcelAddress()
        {
            var token = new Token("A1", TokenType.ExcelAddress);
            var expression = _factory.Create(token);
            Assert.IsInstanceOfType(expression, typeof(ExcelAddressExpression));
        }

        [TestMethod]
        public void ShouldReturnNamedValueExpressionWhenTokenIsNamedValue()
        {
            var token = new Token("NamedValue", TokenType.NameValue);
            var expression = _factory.Create(token);
            Assert.IsInstanceOfType(expression, typeof(NamedValueExpression));
        }
    }
}
