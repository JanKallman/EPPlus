using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class TokenFactoryTests
    {
        private ITokenFactory _tokenFactory;
        private INameValueProvider _nameValueProvider;


        [TestInitialize]
        public void Setup()
        {
            var context = ParsingContext.Create();
            var excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _nameValueProvider = MockRepository.GenerateStub<INameValueProvider>();
            _tokenFactory = new TokenFactory(context.Configuration.FunctionRepository, _nameValueProvider);
        }

        [TestCleanup]
        public void Cleanup()
        {
      
        }

        [TestMethod]
        public void ShouldCreateAStringToken()
        {
            var input = "'";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("'", token.Value);
            Assert.AreEqual(TokenType.String, token.TokenType);
        }

        [TestMethod]
        public void ShouldCreatePlusAsOperatorToken()
        {
            var input = "+";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("+", token.Value);
            Assert.AreEqual(TokenType.Operator, token.TokenType);
        }

        [TestMethod]
        public void ShouldCreateMinusAsOperatorToken()
        {
            var input = "-";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("-", token.Value);
            Assert.AreEqual(TokenType.Operator, token.TokenType);
        }

        [TestMethod]
        public void ShouldCreateMultiplyAsOperatorToken()
        {
            var input = "*";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("*", token.Value);
            Assert.AreEqual(TokenType.Operator, token.TokenType);
        }

        [TestMethod]
        public void ShouldCreateDivideAsOperatorToken()
        {
            var input = "/";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("/", token.Value);
            Assert.AreEqual(TokenType.Operator, token.TokenType);
        }

        [TestMethod]
        public void ShouldCreateEqualsAsOperatorToken()
        {
            var input = "=";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("=", token.Value);
            Assert.AreEqual(TokenType.Operator, token.TokenType);
        }

        [TestMethod]
        public void ShouldCreateIntegerAsIntegerToken()
        {
            var input = "23";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("23", token.Value);
            Assert.AreEqual(TokenType.Integer, token.TokenType);
        }

        [TestMethod]
        public void ShouldCreateBooleanAsBooleanToken()
        {
            var input = "true";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("true", token.Value);
            Assert.AreEqual(TokenType.Boolean, token.TokenType);
        }

        [TestMethod]
        public void ShouldCreateDecimalAsDecimalToken()
        {
            var input = "23.3";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.AreEqual("23.3", token.Value);
            Assert.AreEqual(TokenType.Decimal, token.TokenType);
        }

        [TestMethod]
        public void CreateShouldReadFunctionsFromFuncRepository()
        {
            var input = "Text";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
            Assert.AreEqual(TokenType.Function, token.TokenType);
            Assert.AreEqual("Text", token.Value);
        }

        [TestMethod]
        public void CreateShouldCreateExcelAddressAsExcelAddressToken()
        {
            var input = "A1";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
            Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
            Assert.AreEqual("A1", token.Value);
        }

        [TestMethod]
        public void CreateShouldCreateExcelRangeAsExcelAddressToken()
        {
            var input = "A1:B15";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
            Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
            Assert.AreEqual("A1:B15", token.Value);
        }

        [TestMethod]
        public void CreateShouldCreateExcelRangeOnOtherSheetAsExcelAddressToken()
        {
            var input = "ws!A1:B15";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
            Assert.AreEqual(TokenType.ExcelAddress, token.TokenType);
            Assert.AreEqual("WS!A1:B15", token.Value);
        }

        [TestMethod]
        public void CreateShouldCreateNamedValueAsExcelAddressToken()
        {
            var input = "NamedValue";
            _nameValueProvider.Stub(x => x.IsNamedValue("NamedValue","")).Return(true);
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
            Assert.AreEqual(TokenType.NameValue, token.TokenType);
            Assert.AreEqual("NamedValue", token.Value);
        }
    }
}
