using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using FakeItEasy;

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class CellReferenceProviderTests
    {
        private ExcelDataProvider _provider;

        [TestInitialize]
        public void Setup()
        {
            _provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => _provider.ExcelMaxRows).Returns(5000);
        }

        [TestMethod]
        public void ShouldReturnReferencedSingleAddress()
        {
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);
            parsingContext.Configuration.SetLexer(new Lexer(parsingContext.Configuration.FunctionRepository, parsingContext.NameValueProvider));
            parsingContext.RangeAddressFactory = new RangeAddressFactory(_provider);
            var provider = new CellReferenceProvider();
            var result = provider.GetReferencedAddresses("A1", parsingContext);
            Assert.AreEqual("A1", result.First());
        }

        [TestMethod]
        public void ShouldReturnReferencedMultipleAddresses()
        {
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);
            parsingContext.Configuration.SetLexer(new Lexer(parsingContext.Configuration.FunctionRepository, parsingContext.NameValueProvider));
            parsingContext.RangeAddressFactory = new RangeAddressFactory(_provider);
            var provider = new CellReferenceProvider();
            var result = provider.GetReferencedAddresses("A1:A2", parsingContext);
            Assert.AreEqual("A1", result.First());
            Assert.AreEqual("A2", result.Last());
        }
    }
}
