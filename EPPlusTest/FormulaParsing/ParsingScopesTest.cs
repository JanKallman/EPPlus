using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using Rhino.Mocks;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class ParsingScopesTest
    {
        private ParsingScopes _parsingScopes;
        private IParsingLifetimeEventHandler _lifeTimeEventHandler;

        [TestInitialize]
        public void Setup()
        {
            _lifeTimeEventHandler = MockRepository.GenerateStub<IParsingLifetimeEventHandler>();
            _parsingScopes = new ParsingScopes(_lifeTimeEventHandler);
        }

        [TestMethod]
        public void CreatedScopeShouldBeCurrentScope()
        {
            using (var scope = _parsingScopes.NewScope(RangeAddress.Empty))
            {
                Assert.AreEqual(_parsingScopes.Current, scope);
            }
        }

        [TestMethod]
        public void CurrentScopeShouldHandleNestedScopes()
        {
            using (var scope1 = _parsingScopes.NewScope(RangeAddress.Empty))
            {
                Assert.AreEqual(_parsingScopes.Current, scope1);
                using (var scope2 = _parsingScopes.NewScope(RangeAddress.Empty))
                {
                    Assert.AreEqual(_parsingScopes.Current, scope2);
                }
                Assert.AreEqual(_parsingScopes.Current, scope1);
            }
            Assert.IsNull(_parsingScopes.Current);
        }

        [TestMethod]
        public void CurrentScopeShouldBeNullWhenScopeHasTerminated()
        {
            using (var scope = _parsingScopes.NewScope(RangeAddress.Empty))
            { }
            Assert.IsNull(_parsingScopes.Current);
        }

        [TestMethod]
        public void NewScopeShouldSetParentOnCreatedScopeIfParentScopeExisted()
        {
            using (var scope1 = _parsingScopes.NewScope(RangeAddress.Empty))
            {
                using (var scope2 = _parsingScopes.NewScope(RangeAddress.Empty))
                {
                    Assert.AreEqual(scope1, scope2.Parent);
                }
            }
        }

        [TestMethod]
        public void LifetimeEventHandlerShouldBeCalled()
        {
            using (var scope = _parsingScopes.NewScope(RangeAddress.Empty))
            { }
            _lifeTimeEventHandler.AssertWasCalled(x => x.ParsingCompleted());
        }
    }
}