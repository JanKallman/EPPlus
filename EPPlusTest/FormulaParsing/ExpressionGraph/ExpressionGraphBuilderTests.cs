using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class ExpressionGraphBuilderTests
    {
        private IExpressionGraphBuilder _graphBuilder;

        [TestInitialize]
        public void Setup()
        {
            var excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            var parsingContext = ParsingContext.Create();
            _graphBuilder = new ExpressionGraphBuilder(excelDataProvider, parsingContext);
        }

        [TestCleanup]
        public void Cleanup()
        {

        }

        [TestMethod]
        public void BuildShouldNotUseStringIdentifyersWhenBuildingStringExpression()
        {
            var tokens = new List<Token>
            {
                new Token("'", TokenType.String),
                new Token("abc", TokenType.StringContent),
                new Token("'", TokenType.String)
            };

            var result = _graphBuilder.Build(tokens);

            Assert.AreEqual(1, result.Expressions.Count());
        }

        [TestMethod]
        public void BuildShouldNotEvaluateExpressionsWithinAString()
        {
            var tokens = new List<Token>
            {
                new Token("'", TokenType.String),
                new Token("1 + 2", TokenType.StringContent),
                new Token("'", TokenType.String)
            };

            var result = _graphBuilder.Build(tokens);

            Assert.AreEqual("1 + 2", result.Expressions.First().Compile().Result);
        }

        [TestMethod]
        public void BuildShouldSetOperatorOnGroupExpressionCorrectly()
        {
            var tokens = new List<Token>
            {
                new Token("(", TokenType.OpeningBracket),
                new Token("2", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("4", TokenType.Integer),
                new Token(")", TokenType.ClosingBracket),
                new Token("*", TokenType.Operator),
                new Token("2", TokenType.Integer)
            };
            var result = _graphBuilder.Build(tokens);

            Assert.AreEqual(Operator.Multiply.Operator, result.Expressions.First().Operator.Operator);

        }

        [TestMethod]
        public void BuildShouldSetChildrenOnGroupExpression()
        {
            var tokens = new List<Token>
            {
                new Token("(", TokenType.OpeningBracket),
                new Token("2", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("4", TokenType.Integer),
                new Token(")", TokenType.ClosingBracket),
                new Token("*", TokenType.Operator),
                new Token("2", TokenType.Integer)
            };
            var result = _graphBuilder.Build(tokens);

            Assert.IsInstanceOfType(result.Expressions.First(), typeof(GroupExpression));
            Assert.AreEqual(2, result.Expressions.First().Children.Count());
        }

        [TestMethod]
        public void BuildShouldSetNextOnGroupedExpression()
        {
            var tokens = new List<Token>
            {
                new Token("(", TokenType.OpeningBracket),
                new Token("2", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("4", TokenType.Integer),
                new Token(")", TokenType.ClosingBracket),
                new Token("*", TokenType.Operator),
                new Token("2", TokenType.Integer)
            };
            var result = _graphBuilder.Build(tokens);

            Assert.IsNotNull(result.Expressions.First().Next);
            Assert.IsInstanceOfType(result.Expressions.First().Next, typeof(IntegerExpression));

        }

        [TestMethod]
        public void BuildShouldBuildFunctionExpressionIfFirstTokenIsFunction()
        {
            var tokens = new List<Token>
            {
                new Token("CStr", TokenType.Function),
                new Token("(", TokenType.OpeningBracket),
                new Token("2", TokenType.Integer),
                new Token(")", TokenType.ClosingBracket),
            };
            var result = _graphBuilder.Build(tokens);

            Assert.AreEqual(1, result.Expressions.Count());
            Assert.IsInstanceOfType(result.Expressions.First(), typeof(FunctionExpression));
        }

        [TestMethod]
        public void BuildShouldSetChildrenOnFunctionExpression()
        {
            var tokens = new List<Token>
            {
                new Token("CStr", TokenType.Function),
                new Token("(", TokenType.OpeningBracket),
                new Token("2", TokenType.Integer),
                new Token(")", TokenType.ClosingBracket)
            };
            var result = _graphBuilder.Build(tokens);

            Assert.AreEqual(1, result.Expressions.First().Children.Count());
            Assert.IsInstanceOfType(result.Expressions.First().Children.First(), typeof(GroupExpression));
            Assert.IsInstanceOfType(result.Expressions.First().Children.First().Children.First(), typeof(IntegerExpression));
            Assert.AreEqual(2d, result.Expressions.First().Children.First().Compile().Result);
        }

        [TestMethod]
        public void BuildShouldAddOperatorToFunctionExpression()
        {
            var tokens = new List<Token>
            {
                new Token("CStr", TokenType.Function),
                new Token("(", TokenType.OpeningBracket),
                new Token("2", TokenType.Integer),
                new Token(")", TokenType.ClosingBracket),
                new Token("&", TokenType.Operator),
                new Token("A", TokenType.StringContent)
            };
            var result = _graphBuilder.Build(tokens);

            Assert.AreEqual(1, result.Expressions.First().Children.Count());
            Assert.AreEqual(2, result.Expressions.Count());
        }

        [TestMethod]
        public void BuildShouldAddCommaSeparatedFunctionArgumentsAsChildrenToFunctionExpression()
        {
            var tokens = new List<Token>
            {
                new Token("Text", TokenType.Function),
                new Token("(", TokenType.OpeningBracket),
                new Token("2", TokenType.Integer),
                new Token(",", TokenType.Comma),
                new Token("3", TokenType.Integer),
                new Token(")", TokenType.ClosingBracket),
                new Token("&", TokenType.Operator),
                new Token("A", TokenType.StringContent)
            };

            var result = _graphBuilder.Build(tokens);

            Assert.AreEqual(2, result.Expressions.First().Children.Count());
        }

        [TestMethod]
        public void BuildShouldCreateASingleExpressionOutOfANegatorAndANumericToken()
        {
            var tokens = new List<Token>
            {
                new Token("-", TokenType.Negator),
                new Token("2", TokenType.Integer),
            };

            var result = _graphBuilder.Build(tokens);

            Assert.AreEqual(1, result.Expressions.Count());
            Assert.AreEqual(-2d, result.Expressions.First().Compile().Result);
        }

        [TestMethod]
        public void BuildShouldHandleEnumerableTokens()
        {
            var tokens = new List<Token>
            {
                new Token("Text", TokenType.Function),
                new Token("(", TokenType.OpeningBracket),
                new Token("{", TokenType.OpeningEnumerable),
                new Token("2", TokenType.Integer),
                new Token(",", TokenType.Comma),
                new Token("3", TokenType.Integer),
                new Token("}", TokenType.ClosingEnumerable),
                new Token(")", TokenType.ClosingBracket)
            };

            var result = _graphBuilder.Build(tokens);
            var funcArgExpression = result.Expressions.First().Children.First();
            Assert.IsInstanceOfType(funcArgExpression, typeof(FunctionArgumentExpression));

            var enumerableExpression = funcArgExpression.Children.First();

            Assert.IsInstanceOfType(enumerableExpression, typeof(EnumerableExpression));
            Assert.AreEqual(2, enumerableExpression.Children.Count(), "Enumerable.Count was not 2");
        }
    }
}
