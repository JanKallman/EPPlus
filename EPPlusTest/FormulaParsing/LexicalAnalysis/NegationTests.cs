using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing;


namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class NegationTests
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
        public void ShouldSetNegatorOnFirstTokenIfFirstCharIsMinus()
        {
            var input = "-1";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(2, tokens.Count());
            Assert.AreEqual(TokenType.Negator, tokens.First().TokenType);
        }

        [TestMethod]
        public void ShouldChangePlusToMinusIfNegatorIsPresent()
        {
            var input = "1 + -1";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(3, tokens.Count());
            Assert.AreEqual(TokenType.Operator, tokens.ElementAt(1).TokenType);
            Assert.AreEqual("-", tokens.ElementAt(1).Value);
        }

        [TestMethod]
        public void ShouldSetNegatorOnTokenInsideParenthethis()
        {
            var input = "1 + (-1 * 2)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(8, tokens.Count());
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(3).TokenType);
        }

        [TestMethod]
        public void ShouldSetNegatorOnTokenInsideFunctionCall()
        {
            var input = "Ceiling(-1, -0.1)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(8, tokens.Count());
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(2).TokenType);
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(5).TokenType, "Negator after comma was not identified");
        }

        [TestMethod]
        public void ShouldSetNegatorOnTokenInEnumerable()
        {
            var input = "{-1}";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(1).TokenType);
        }

        [TestMethod]
        public void ShouldSetNegatorOnExcelAddress()
        {
            var input = "-A1";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(TokenType.Negator, tokens.ElementAt(0).TokenType);
            Assert.AreEqual(TokenType.ExcelAddress, tokens.ElementAt(1).TokenType);
        }
    }
}
