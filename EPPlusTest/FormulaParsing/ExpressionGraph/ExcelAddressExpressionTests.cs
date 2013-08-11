using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class ExcelAddressExpressionTests
    {
        private ParsingContext _parsingContext;
        private ParsingScope _scope;

        private ExcelCell CreateItem(object val)
        {
            return new ExcelCell(val, null, 0, 0);
        }

        [TestInitialize]
        public void Setup()
        {
            _parsingContext = ParsingContext.Create();
            _scope = _parsingContext.Scopes.NewScope(RangeAddress.Empty);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _scope.Dispose();
        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ConstructorShouldThrowIfExcelDataProviderIsNull()
        {
            new ExcelAddressExpression("A1", null, _parsingContext);
        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ConstructorShouldThrowIfParsingContextIsNull()
        {
            new ExcelAddressExpression("A1", MockRepository.GenerateStub<ExcelDataProvider>(), null);
        }

        [TestMethod]
        public void ShouldCallReturnResultFromProvider()
        {
            var expectedAddress = "A1";
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            provider
                .Stub(x => x.GetRangeValues(string.Empty, expectedAddress))
                .Return(new object[]{ 1 });

            var expression = new ExcelAddressExpression(expectedAddress, provider, _parsingContext);
            var result = expression.Compile();
            Assert.AreEqual(1, result.Result);
        }

        [TestMethod]
        public void CompileShouldReturnAddress()
        {
            var expectedAddress = "A1";
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            provider
                .Stub(x => x.GetRangeValues(expectedAddress))
                .Return(new ExcelCell[] { CreateItem(1) });

            var expression = new ExcelAddressExpression(expectedAddress, provider, _parsingContext);
            expression.ParentIsLookupFunction = true;
            var result = expression.Compile();
            Assert.AreEqual(expectedAddress, result.Result);

        }
    }
}
