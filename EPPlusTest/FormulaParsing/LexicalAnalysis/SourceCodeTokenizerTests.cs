using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class SourceCodeTokenizerTests
    {
        private SourceCodeTokenizer _tokenizer;

        [TestInitialize]
        public void Setup()
        {
            var context = ParsingContext.Create();
            _tokenizer = new SourceCodeTokenizer(context.Configuration.FunctionRepository, null);
        }

        [TestCleanup]
        public void Cleanup()
        {
        }

        [TestMethod]
        public void ShouldCreateTokensForStringCorrectly()
        {
            var input = "'abc123'";
            var tokens = _tokenizer.Tokenize(input);
            
            Assert.AreEqual(3, tokens.Count());
            Assert.AreEqual(TokenType.String, tokens.First().TokenType);
            Assert.AreEqual(TokenType.StringContent, tokens.ElementAt(1).TokenType);
            Assert.AreEqual(TokenType.String, tokens.Last().TokenType);
        }

        [TestMethod]
        public void ShouldCreateTokensForStringCorrectly2()
        {
            var input = "\"abc123\"";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(3, tokens.Count());
            Assert.AreEqual(TokenType.String, tokens.First().TokenType);
            Assert.AreEqual(TokenType.StringContent, tokens.ElementAt(1).TokenType);
            Assert.AreEqual(TokenType.String, tokens.Last().TokenType);
        }

        [TestMethod]
        public void ShouldIgnoreTokenSeparatorsInAString()
        {
            var input = "'ab(c)d'";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(3, tokens.Count());
        }

        [TestMethod]
        public void ShouldCreateTokensForFunctionCorrectly()
        {
            var input = "Text(2)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(4, tokens.Count());
            Assert.AreEqual(TokenType.Function, tokens.First().TokenType);
            Assert.AreEqual(TokenType.OpeningParenthesis, tokens.ElementAt(1).TokenType);
            Assert.AreEqual(TokenType.Integer, tokens.ElementAt(2).TokenType);
            Assert.AreEqual("2", tokens.ElementAt(2).Value);
            Assert.AreEqual(TokenType.ClosingParenthesis, tokens.Last().TokenType);
        }

        [TestMethod]
        public void ShouldHandleMultipleCharOperatorCorrectly()
        {
            var input = "1 <= 2";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(3, tokens.Count());
            Assert.AreEqual("<=", tokens.ElementAt(1).Value);
            Assert.AreEqual(TokenType.Operator, tokens.ElementAt(1).TokenType);
        }

        [TestMethod]
        public void ShouldCreateTokensForEnumerableCorrectly()
        {
            var input = "Text({1;2})";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(8, tokens.Count());
            Assert.AreEqual(TokenType.OpeningEnumerable, tokens.ElementAt(2).TokenType);
            Assert.AreEqual(TokenType.ClosingEnumerable, tokens.ElementAt(6).TokenType);
        }

        [TestMethod]
        public void ShouldCreateTokensForExcelAddressCorrectly()
        {
            var input = "Text(A1)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(TokenType.ExcelAddress, tokens.ElementAt(2).TokenType);
        }

        [TestMethod]
        public void ShouldCreateTokenForPercentAfterDecimal()
        {
            var input = "1,23%";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(TokenType.Percent, tokens.Last().TokenType);
        }

        [TestMethod]
        public void ShouldHandleEmptyString()
        {
            var input = "IF(I10>=0;IF(O10>I10;((O10-I10)*$B10)/$C$27;IF(O10<0;(O10*$B10)/$C$27;\"\"));IF(O10<0;((O10-I10)*$B10)/$C$27;IF(O10>0;(O10*$B10)/$C$27;)))";
            //var input = "2<=3";
            var tokens = _tokenizer.Tokenize(input);
            //var factory = new ExpressionGraphBuilder()
            Assert.Fail();
        }
    }
}
