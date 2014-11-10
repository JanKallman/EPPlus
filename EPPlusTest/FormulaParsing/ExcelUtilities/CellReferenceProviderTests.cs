using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class CellReferenceProviderTests
    {
        private ExcelDataProvider _provider;

        [TestInitialize]
        public void Setup()
        {
            _provider = MockRepository.GenerateStub<ExcelDataProvider>();
            _provider.Stub(x => x.ExcelMaxRows).Return(5000);
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
