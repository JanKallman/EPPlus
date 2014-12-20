using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class RangeAddressTests
    {
        private RangeAddressFactory _factory;

        [TestInitialize]
        public void Setup()
        {
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            _factory = new RangeAddressFactory(provider);
        }

        [TestMethod]
        public void CollideShouldReturnTrueIfRangesCollides()
        {
            var address1 = _factory.Create("A1:A6");
            var address2 = _factory.Create("A5");
            Assert.IsTrue(address1.CollidesWith(address2));
        }

        [TestMethod]
        public void CollideShouldReturnFalseIfRangesDoesNotCollide()
        {
            var address1 = _factory.Create("A1:A6");
            var address2 = _factory.Create("A8");
            Assert.IsFalse(address1.CollidesWith(address2));
        }

        [TestMethod]
        public void CollideShouldReturnFalseIfRangesCollidesButWorksheetNameDiffers()
        {
            var address1 = _factory.Create("Ws!A1:A6");
            var address2 = _factory.Create("A5");
            Assert.IsFalse(address1.CollidesWith(address2));
        }
    }
}
