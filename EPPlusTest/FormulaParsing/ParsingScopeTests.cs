using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class ParsingScopeTests
    {
        private IParsingLifetimeEventHandler _lifeTimeEventHandler;
        private ParsingScopes _parsingScopes;
        private RangeAddressFactory _factory;

        [TestInitialize]
        public void Setup()
        {
            var provider = A.Fake<ExcelDataProvider>();
            _factory = new RangeAddressFactory(provider);
            _lifeTimeEventHandler = A.Fake<IParsingLifetimeEventHandler>();
            _parsingScopes = A.Fake<ParsingScopes>();
        }

        [TestMethod]
        public void ConstructorShouldSetAddress()
        {
            var expectedAddress =  _factory.Create("A1");
            var scope = new ParsingScope(_parsingScopes, expectedAddress);
            Assert.AreEqual(expectedAddress, scope.Address);
        }

        [TestMethod]
        public void ConstructorShouldSetParent()
        {
            var parent = new ParsingScope(_parsingScopes, _factory.Create("A1"));
            var scope = new ParsingScope(_parsingScopes, parent, _factory.Create("A2"));
            Assert.AreEqual(parent, scope.Parent);
        }

        [TestMethod]
        public void ScopeShouldCallKillScopeOnDispose()
        {
            var scope = new ParsingScope(_parsingScopes, _factory.Create("A1"));
            ((IDisposable)scope).Dispose();
           A.CallTo(() => _parsingScopes.KillScope(scope)).MustHaveHappened();
        }
    }
}
