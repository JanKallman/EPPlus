using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
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
            var input = "\"abc123\"";
            var tokens = _tokenizer.Tokenize(input);

            Assert.AreEqual(3, tokens.Count());
            Assert.AreEqual(TokenType.String, tokens.First().TokenType);
            Assert.AreEqual(TokenType.StringContent, tokens.ElementAt(1).TokenType);
            Assert.AreEqual(TokenType.String, tokens.Last().TokenType);
        }

        [TestMethod]
        public void ShouldTokenizeStringCorrectly()
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
        public void ShouldIgnoreTwoSubsequentStringIdentifyers()
        {
            var input = "\"hello\"\"world\"";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(3, tokens.Count());
            Assert.AreEqual("hello\"world", tokens.ElementAt(1).Value);
        }

        [TestMethod]
        public void ShouldIgnoreTwoSubsequentStringIdentifyers2()
        {
            //using (var pck = new ExcelPackage(new FileInfo("c:\\temp\\QuoteIssue2.xlsx")))
            //{
            //    pck.Workbook.Worksheets.First().Calculate();
            //}
            var input = "\"\"\"\"\"\"";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(TokenType.StringContent, tokens.ElementAt(1).TokenType);
        }

        [TestMethod]
        public void TokenizerShouldIgnoreOperatorInString()
        {
            var input = "\"*\"";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(TokenType.StringContent, tokens.ElementAt(1).TokenType);
        }

        [TestMethod]
        public void TokenizerShouldHandleWorksheetNameWithMinus()
        {
            var input = "'A-B'!A1";
            var tokens = _tokenizer.Tokenize(input);
            Assert.AreEqual(1, tokens.Count());
        }

        [TestMethod]
        public void TestBug9_12_14()
        {
            //(( W60 -(- W63 )-( W29 + W30 + W31 ))/( W23 + W28 + W42 - W51 )* W4 )
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("test");
                for (var x = 1; x <= 10; x++)
                {
                    ws1.Cells[x, 1].Value = x;
                }

                ws1.Cells["A11"].Formula = "(( A1 -(- A2 )-( A3 + A4 + A5 ))/( A6 + A7 + A8 - A9 )* A5 )";
                //ws1.Cells["A11"].Formula = "(-A2 + 1 )";
                ws1.Calculate();
                var result = ws1.Cells["A11"].Value;
                Assert.AreEqual(-3.75, result);
            }
        }
    }
}
