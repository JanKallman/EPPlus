using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class TokenHandlerTests
    {
        private TokenizerContext _tokenizerContext;
        private TokenHandler _handler;

        [TestInitialize]
        public void Init()
        {
            _tokenizerContext = new TokenizerContext("test");
            InitHandler(_tokenizerContext);
        }

        private void InitHandler(TokenizerContext context)
        {
            var parsingContext = ParsingContext.Create();
            var tokenFactory = new TokenFactory(parsingContext.Configuration.FunctionRepository, null);
            _handler = new TokenHandler(_tokenizerContext, tokenFactory, new TokenSeparatorProvider()); 
        }

        [TestMethod]
        public void HasMoreTokensShouldBeTrueWhenTokensExists()
        {
            Assert.IsTrue(_handler.HasMore());
        }

        [TestMethod]
        public void HasMoreTokensShouldBeFalseWhenAllAreHandled()
        {
            for (var x = 0; x < "test".Length; x++ )
            {
                _handler.Next();
            }
            Assert.IsFalse(_handler.HasMore());
        }
    }
}
