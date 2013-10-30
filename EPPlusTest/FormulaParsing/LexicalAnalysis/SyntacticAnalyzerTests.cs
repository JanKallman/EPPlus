using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class SyntacticAnalyzerTests
    {
        private ISyntacticAnalyzer _analyser;

        [TestInitialize]
        public void Setup()
        {
            _analyser = new SyntacticAnalyzer();
        }

        [TestMethod]
        public void ShouldPassIfParenthesisAreWellformed()
        {
            var input = new List<Token>
            {
                new Token("(", TokenType.OpeningParenthesis),
                new Token("1", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("2", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis)
            };
            _analyser.Analyze(input);
        }

        [TestMethod, ExpectedException(typeof(FormatException))]
        public void ShouldThrowExceptionIfParenthesesAreNotWellformed()
        {
            var input = new List<Token>
            {
                new Token("(", TokenType.OpeningParenthesis),
                new Token("1", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("2", TokenType.Integer)
            };
            _analyser.Analyze(input);
        }

        [TestMethod]
        public void ShouldPassIfStringIsWellformed()
        {
            var input = new List<Token>
            {
                new Token("'", TokenType.String),
                new Token("abc123", TokenType.StringContent),
                new Token("'", TokenType.String)
            };
            _analyser.Analyze(input);
        }

        [TestMethod, ExpectedException(typeof(FormatException))]
        public void ShouldThrowExceptionIfStringHasNotClosing()
        {
            var input = new List<Token>
            {
                new Token("'", TokenType.String),
                new Token("abc123", TokenType.StringContent)
            };
            _analyser.Analyze(input);
        }


        [TestMethod, ExpectedException(typeof(UnrecognizedTokenException))]
        public void ShouldThrowExceptionIfThereIsAnUnrecognizedToken()
        {
            var input = new List<Token>
            {
                new Token("abc123", TokenType.Unrecognized)
            };
            _analyser.Analyze(input);
        }
    }
}
